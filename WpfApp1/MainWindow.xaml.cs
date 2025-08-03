using NAudio.CoreAudioApi;
using NAudio.Dsp;
using NAudio.Wave;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        #region Windows API for click-through
        [DllImport("user32.dll")]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        public const int GWL_EXSTYLE = -20;
        public const int WS_EX_LAYERED = 0x80000;
        public const int WS_EX_TRANSPARENT = 0x20;
        #endregion

        #region Fields
        private WasapiLoopbackCapture audioCapture;
        private DispatcherTimer renderTimer;
        private DispatcherTimer titleBarTimer;

        // Dynamic settings
        private int FFT_SIZE = 512;
        private int BAR_COUNT = 48; // More bars for better bass representation
        private const double TARGET_FPS = 120.0;

        private Complex[] fftBuffer;
        private float[] frequencyData;
        private float[] previousData;

        private Rectangle[] visualizerBars;
        private Rectangle[] glowBars;

        // Customizable settings
        private double smoothingFactor = 0.25;
        private double bassBoost = 6.0; // Higher bass boost
        private double midBoost = 2.5;
        private double highBoost = 1.5;
        private double sensitivity = 1.0;

        // Color schemes
        private readonly Color[][] colorSchemes = new Color[][]
        {
            // Scheme 1: Bass Heavy (Red to Yellow)
            new Color[] {
                Color.FromRgb(255, 0, 0),      // Deep Red (Bass)
                Color.FromRgb(255, 50, 0),     // Red-Orange
                Color.FromRgb(255, 100, 0),    // Orange
                Color.FromRgb(255, 150, 0),    // Light Orange
                Color.FromRgb(255, 200, 0),    // Yellow-Orange
                Color.FromRgb(255, 255, 0),    // Yellow
                Color.FromRgb(200, 255, 0),    // Yellow-Green
                Color.FromRgb(150, 255, 0)     // Light Green
            },
            
            // Scheme 2: Neon (Purple to Cyan)
            new Color[] {
                Color.FromRgb(128, 0, 255),    // Purple (Bass)
                Color.FromRgb(150, 0, 255),    // Light Purple
                Color.FromRgb(200, 0, 255),    // Pink-Purple
                Color.FromRgb(255, 0, 200),    // Pink
                Color.FromRgb(255, 0, 150),    // Hot Pink
                Color.FromRgb(255, 0, 100),    // Red-Pink
                Color.FromRgb(0, 150, 255),    // Blue
                Color.FromRgb(0, 255, 255)     // Cyan
            },
            
            // Scheme 3: Fire (Deep Red to White)
            new Color[] {
                Color.FromRgb(139, 0, 0),      // Dark Red (Bass)
                Color.FromRgb(178, 34, 34),    // Fire Brick
                Color.FromRgb(220, 20, 60),    // Crimson
                Color.FromRgb(255, 69, 0),     // Red-Orange
                Color.FromRgb(255, 140, 0),    // Dark Orange
                Color.FromRgb(255, 165, 0),    // Orange
                Color.FromRgb(255, 215, 0),    // Gold
                Color.FromRgb(255, 255, 255)   // White
            },
            
            // Scheme 4: Ocean (Deep Blue to Cyan)
            new Color[] {
                Color.FromRgb(0, 0, 139),      // Dark Blue (Bass)
                Color.FromRgb(0, 100, 200),    // Blue
                Color.FromRgb(0, 150, 255),    // Light Blue
                Color.FromRgb(0, 200, 255),    // Sky Blue
                Color.FromRgb(0, 255, 200),    // Aqua
                Color.FromRgb(0, 255, 150),    // Light Aqua
                Color.FromRgb(100, 255, 200),  // Mint
                Color.FromRgb(200, 255, 255)   // Light Cyan
            }
        };

        private int currentColorScheme = 0;
        private bool isClickThrough = true;
        private bool showSettings = false;
        private bool showTitleBar = false;
        #endregion

        public MainWindow()
        {
            InitializeComponent();
            InitializeVisualizer();
            SetupEventHandlers();
            CreateContextMenu();
        }

        #region Initialization
        private void InitializeVisualizer()
        {
            // Window setup for true transparency overlay
            this.Topmost = true;
            this.WindowState = WindowState.Maximized;
            this.WindowStyle = WindowStyle.None;
            this.AllowsTransparency = true;
            this.Background = Brushes.Transparent;
            this.ShowInTaskbar = false;

            // Initialize arrays
            InitializeArrays();

            // Create visualizer elements
            CreateVisualizerBars();

            // Initialize settings text immediately
            UpdateSettingsText();

            // Show settings on startup as requested
            showSettings = true;
            SettingsPanel.Visibility = Visibility.Visible;

            // Auto-hide settings after 5 seconds
            var settingsTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(5)
            };
            settingsTimer.Tick += (s, e) =>
            {
                if (showSettings)
                {
                    ToggleSettings();
                }
                settingsTimer.Stop();
            };
            settingsTimer.Start();
        }

        private void InitializeArrays()
        {
            fftBuffer = new Complex[FFT_SIZE];
            frequencyData = new float[BAR_COUNT];
            previousData = new float[BAR_COUNT];
            visualizerBars = new Rectangle[BAR_COUNT];
            glowBars = new Rectangle[BAR_COUNT];
        }

        private void CreateVisualizerBars()
        {
            // Clear existing bars if recreating
            VisualizerCanvas.Children.Clear();

            for (int i = 0; i < BAR_COUNT; i++)
            {
                // Create glow bars
                glowBars[i] = new Rectangle
                {
                    RadiusX = 2,
                    RadiusY = 2,
                    Opacity = 0.5
                };

                // Create main bars
                visualizerBars[i] = new Rectangle
                {
                    RadiusX = 1,
                    RadiusY = 1
                };

                VisualizerCanvas.Children.Add(glowBars[i]);
                VisualizerCanvas.Children.Add(visualizerBars[i]);
            }
        }

        private void CreateContextMenu()
        {
            ContextMenu contextMenu = new ContextMenu();

            // Settings
            MenuItem settingsItem = new MenuItem();
            settingsItem.Header = "⚙ Settings";
            settingsItem.Click += (s, e) => ToggleSettings();
            contextMenu.Items.Add(settingsItem);

            contextMenu.Items.Add(new Separator());

            // Title Bar
            MenuItem titleBarItem = new MenuItem();
            titleBarItem.Header = "📊 Show Title Bar";
            titleBarItem.Click += (s, e) => ToggleTitleBar();
            contextMenu.Items.Add(titleBarItem);

            // Click Through
            MenuItem clickThroughItem = new MenuItem();
            clickThroughItem.Header = "👻 Toggle Click-Through";
            clickThroughItem.Click += (s, e) => ToggleClickThrough();
            contextMenu.Items.Add(clickThroughItem);

            contextMenu.Items.Add(new Separator());

            // Color Schemes
            MenuItem colorItem = new MenuItem();
            colorItem.Header = "🎨 Change Colors";
            colorItem.Click += (s, e) => CycleColorScheme();
            contextMenu.Items.Add(colorItem);

            // Restart Audio
            MenuItem restartItem = new MenuItem();
            restartItem.Header = "🔄 Restart Audio";
            restartItem.Click += (s, e) => RestartAudioCapture();
            contextMenu.Items.Add(restartItem);

            contextMenu.Items.Add(new Separator());

            // Minimize
            MenuItem minimizeItem = new MenuItem();
            minimizeItem.Header = "➖ Minimize";
            minimizeItem.Click += (s, e) => MinimizeWindow();
            contextMenu.Items.Add(minimizeItem);

            // Exit
            MenuItem exitItem = new MenuItem();
            exitItem.Header = "❌ Exit";
            exitItem.Click += (s, e) => Close();
            contextMenu.Items.Add(exitItem);

            this.ContextMenu = contextMenu;
        }

        private void SetupEventHandlers()
        {
            this.KeyDown += OnKeyDown;
            this.Loaded += OnWindowLoaded;
            this.Closing += OnWindowClosing;
            this.SourceInitialized += OnSourceInitialized;
            this.MouseEnter += OnMouseEnter;
            this.MouseLeave += OnMouseLeave;

            // Add mouse click handler to ensure focus
            this.MouseDown += (s, e) => this.Focus();
        }

        private void OnSourceInitialized(object sender, EventArgs e)
        {
            // Make window click-through initially
            ToggleClickThrough();
        }

        private void OnMouseEnter(object sender, MouseEventArgs e)
        {
            if (!isClickThrough)
            {
                ShowTitleBar();
            }
        }

        private void OnMouseLeave(object sender, MouseEventArgs e)
        {
            HideTitleBarDelayed();
        }

        private void ShowTitleBar()
        {
            showTitleBar = true;
            TitleBar.Visibility = Visibility.Visible;

            // Cancel any pending hide operation
            titleBarTimer?.Stop();
        }

        private void HideTitleBarDelayed()
        {
            titleBarTimer?.Stop();
            titleBarTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(2)
            };
            titleBarTimer.Tick += (s, e) =>
            {
                showTitleBar = false;
                TitleBar.Visibility = Visibility.Collapsed;
                titleBarTimer.Stop();
            };
            titleBarTimer.Start();
        }

        private void ToggleTitleBar()
        {
            showTitleBar = !showTitleBar;
            TitleBar.Visibility = showTitleBar ? Visibility.Visible : Visibility.Collapsed;

            if (showTitleBar)
            {
                titleBarTimer?.Stop(); // Stop auto-hide when manually shown
            }
        }
        #endregion

        #region Window Control Event Handlers
        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            ToggleSettings();
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            MinimizeWindow();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void MinimizeWindow()
        {
            this.WindowState = WindowState.Minimized;
            this.ShowInTaskbar = true; // Show in taskbar when minimized
        }
        #endregion

        #region Audio Processing
        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            // Ensure window has focus to receive key events
            this.Focus();
            this.Activate();

            StartAudioCapture();
            StartRenderLoop();
        }

        private void StartAudioCapture()
        {
            try
            {
                audioCapture = new WasapiLoopbackCapture();
                audioCapture.ShareMode = AudioClientShareMode.Shared;

                audioCapture.DataAvailable += OnAudioDataAvailable;
                audioCapture.RecordingStopped += OnRecordingStopped;
                audioCapture.StartRecording();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to start audio capture: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnAudioDataAvailable(object sender, WaveInEventArgs e)
        {
            ProcessAudioDataDirectly(e.Buffer, e.BytesRecorded);
        }

        private void ProcessAudioDataDirectly(byte[] buffer, int bytesRecorded)
        {
            if (bytesRecorded < FFT_SIZE * 4) return;

            try
            {
                int startOffset = Math.Max(0, bytesRecorded - (FFT_SIZE * 4));

                for (int i = 0; i < FFT_SIZE; i++)
                {
                    int sampleIndex = startOffset + (i * 4);
                    if (sampleIndex + 3 < bytesRecorded)
                    {
                        float sample = BitConverter.ToSingle(buffer, sampleIndex);

                        // Apply Hann windowing
                        double window = 0.5 * (1 - Math.Cos(2 * Math.PI * i / (FFT_SIZE - 1)));
                        fftBuffer[i].X = (float)(sample * window);
                        fftBuffer[i].Y = 0;
                    }
                    else
                    {
                        fftBuffer[i].X = 0;
                        fftBuffer[i].Y = 0;
                    }
                }

                FastFourierTransform.FFT(true, (int)Math.Log2(FFT_SIZE), fftBuffer);
                ConvertToFrequencyBins();
            }
            catch (Exception)
            {
                // Ignore processing errors
            }
        }

        private void ConvertToFrequencyBins()
        {
            // Allocate more bars to bass frequencies (right side)
            int bassBarCount = (int)(BAR_COUNT * 0.4); // 40% for bass
            int midBarCount = (int)(BAR_COUNT * 0.4);  // 40% for mids
            int highBarCount = BAR_COUNT - bassBarCount - midBarCount; // 20% for highs

            int fftHalf = FFT_SIZE / 2;

            // Process bass frequencies (right side bars)
            for (int i = 0; i < bassBarCount; i++)
            {
                int startBin = (int)(i * (fftHalf * 0.1) / bassBarCount); // Use lower 10% of spectrum
                int endBin = Math.Min(startBin + Math.Max(1, (int)(fftHalf * 0.05 / bassBarCount)), (int)(fftHalf * 0.1));

                float sum = CalculateBinSum(startBin, endBin);
                float rawValue = sum * (float)bassBoost * (float)sensitivity;

                int barIndex = BAR_COUNT - 1 - i; // Right side
                frequencyData[barIndex] = (float)(previousData[barIndex] * smoothingFactor + rawValue * (1.0 - smoothingFactor));
                previousData[barIndex] = frequencyData[barIndex];
            }

            // Process mid frequencies (center bars)
            for (int i = 0; i < midBarCount; i++)
            {
                int startBin = (int)(fftHalf * 0.1 + i * (fftHalf * 0.4) / midBarCount);
                int endBin = Math.Min(startBin + Math.Max(1, (int)(fftHalf * 0.4 / midBarCount)), (int)(fftHalf * 0.5));

                float sum = CalculateBinSum(startBin, endBin);
                float rawValue = sum * (float)midBoost * (float)sensitivity;

                int barIndex = bassBarCount + i;
                frequencyData[barIndex] = (float)(previousData[barIndex] * smoothingFactor + rawValue * (1.0 - smoothingFactor));
                previousData[barIndex] = frequencyData[barIndex];
            }

            // Process high frequencies (left side bars)
            for (int i = 0; i < highBarCount; i++)
            {
                int startBin = (int)(fftHalf * 0.5 + i * (fftHalf * 0.5) / highBarCount);
                int endBin = Math.Min(startBin + Math.Max(1, (int)(fftHalf * 0.5 / highBarCount)), fftHalf);

                float sum = CalculateBinSum(startBin, endBin);
                float rawValue = sum * (float)highBoost * (float)sensitivity;

                int barIndex = i;
                frequencyData[barIndex] = (float)(previousData[barIndex] * smoothingFactor + rawValue * (1.0 - smoothingFactor));
                previousData[barIndex] = frequencyData[barIndex];
            }
        }

        private float CalculateBinSum(int startBin, int endBin)
        {
            float sum = 0;
            for (int j = startBin; j < endBin; j++)
            {
                if (j < fftBuffer.Length)
                {
                    float magnitude = (float)Math.Sqrt(fftBuffer[j].X * fftBuffer[j].X + fftBuffer[j].Y * fftBuffer[j].Y);
                    sum += magnitude;
                }
            }
            return sum / Math.Max(1, endBin - startBin);
        }

        private void OnRecordingStopped(object sender, StoppedEventArgs e)
        {
            if (e.Exception != null)
            {
                Dispatcher.BeginInvoke(() =>
                {
                    MessageBox.Show($"Audio capture stopped: {e.Exception.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        private void StartRenderLoop()
        {
            renderTimer = new DispatcherTimer(DispatcherPriority.Render)
            {
                Interval = TimeSpan.FromMilliseconds(1000.0 / TARGET_FPS)
            };
            renderTimer.Tick += OnRenderFrame;
            renderTimer.Start();
        }

        private void OnRenderFrame(object sender, EventArgs e)
        {
            UpdateVisualizer();
        }
        #endregion

        #region Visualization
        private void UpdateVisualizer()
        {
            double canvasWidth = VisualizerCanvas.ActualWidth;
            double canvasHeight = VisualizerCanvas.ActualHeight;

            if (canvasWidth <= 0 || canvasHeight <= 0) return;

            // Position bars at absolute bottom of screen
            double barWidth = canvasWidth * 0.85 / BAR_COUNT;
            double barSpacing = barWidth * 0.1;
            double totalWidth = BAR_COUNT * barWidth + (BAR_COUNT - 1) * barSpacing;
            double startX = (canvasWidth - totalWidth) / 2;
            double maxBarHeight = canvasHeight * 0.6; // Taller bars

            for (int i = 0; i < BAR_COUNT; i++)
            {
                double x = startX + i * (barWidth + barSpacing);

                // Enhanced logarithmic scaling with bass emphasis
                double normalizedHeight = Math.Min(Math.Log10(1 + frequencyData[i] * 80) / 2.2, 1.0);

                // Extra boost for bass bars (right side)
                if (i >= BAR_COUNT - (int)(BAR_COUNT * 0.4))
                {
                    normalizedHeight *= 1.3; // 30% extra height for bass
                }

                double barHeight = Math.Max(4, normalizedHeight * maxBarHeight);
                double y = canvasHeight - barHeight; // Absolute bottom

                // Update bars
                UpdateBarDirect(visualizerBars[i], x, y, barWidth * 0.9, barHeight, i, false);

                double glowHeight = barHeight * 1.2;
                double glowY = y - (glowHeight - barHeight) / 2;
                UpdateBarDirect(glowBars[i], x - 1, glowY, barWidth * 1.1, glowHeight, i, true);
            }
        }

        private void UpdateBarDirect(Rectangle bar, double x, double y, double width, double height, int index, bool isGlow)
        {
            Canvas.SetLeft(bar, x);
            Canvas.SetTop(bar, y);
            bar.Width = width;
            bar.Height = height;

            Color baseColor = GetFrequencyColor(index);

            if (isGlow)
            {
                bar.Fill = new SolidColorBrush(Color.FromArgb(80, baseColor.R, baseColor.G, baseColor.B));
            }
            else
            {
                byte alpha = (byte)Math.Min(255, 180 + frequencyData[index] * 200);
                bar.Fill = new SolidColorBrush(Color.FromArgb(alpha, baseColor.R, baseColor.G, baseColor.B));
            }
        }

        private Color GetFrequencyColor(int barIndex)
        {
            Color[] currentColors = colorSchemes[currentColorScheme];
            double ratio = (double)barIndex / BAR_COUNT;
            int colorIndex = (int)(ratio * currentColors.Length);
            colorIndex = Math.Max(0, Math.Min(colorIndex, currentColors.Length - 1));

            return currentColors[colorIndex];
        }
        #endregion

        #region Settings and Controls
        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            // Debug output to check if key events are being received
            System.Diagnostics.Debug.WriteLine($"Key pressed: {e.Key}");

            switch (e.Key)
            {
                case Key.Escape:
                    Close();
                    break;
                case Key.T:
                    ToggleClickThrough();
                    break;
                case Key.S:
                    ToggleSettings();
                    break;
                case Key.C:
                    CycleColorScheme();
                    break;
                case Key.R:
                    RestartAudioCapture();
                    break;
                case Key.H:
                    ToggleTitleBar();
                    break;
                case Key.M:
                    MinimizeWindow();
                    break;
                case Key.OemPlus:
                case Key.Add:
                    AdjustSensitivity(1.2);
                    break;
                case Key.OemMinus:
                case Key.Subtract:
                    AdjustSensitivity(0.8);
                    break;
                case Key.D1:
                    SetBarCount(24);
                    break;
                case Key.D2:
                    SetBarCount(32);
                    break;
                case Key.D3:
                    SetBarCount(48);
                    break;
                case Key.D4:
                    SetBarCount(64);
                    break;
                case Key.D5:
                    SetBarCount(80);
                    break;
                case Key.B:
                    AdjustBassBoost(1.2);
                    break;
                case Key.V:
                    AdjustBassBoost(0.8);
                    break;
            }

            e.Handled = true; // Mark the event as handled
        }

        private void ToggleSettings()
        {
            showSettings = !showSettings;
            SettingsPanel.Visibility = showSettings ? Visibility.Visible : Visibility.Collapsed;
            UpdateSettingsText();

            // Debug output
            System.Diagnostics.Debug.WriteLine($"Settings toggled: {showSettings}");

            // Ensure window can receive key events when settings are shown
            if (showSettings)
            {
                this.Focus();
                this.Activate();
            }
        }

        private void UpdateSettingsText()
        {
            if (SettingsText != null)
            {
                SettingsText.Text = $"🎵 REAL-TIME AUDIO VISUALIZER SETTINGS 🎵\n\n" +
                                   $"Current Settings:\n" +
                                   $"• Bars: {BAR_COUNT}\n" +
                                   $"• Color Scheme: {currentColorScheme + 1}/4\n" +
                                   $"• Bass Boost: {bassBoost:F1}x\n" +
                                   $"• Sensitivity: {sensitivity:F1}x\n" +
                                   $"• Click-Through: {(isClickThrough ? "ON" : "OFF")}\n" +
                                   $"• Title Bar: {(showTitleBar ? "VISIBLE" : "HIDDEN")}\n\n" +
                                   $"Keyboard Controls:\n" +
                                   $"• ESC: Exit • S: Settings • T: Click-Through\n" +
                                   $"• C: Change Colors • R: Restart Audio\n" +
                                   $"• H: Toggle Title Bar • M: Minimize\n" +
                                   $"• +/-: Sensitivity • B/V: Bass Boost\n" +
                                   $"• 1-5: Bar Count (24/32/48/64/80)\n\n" +
                                   $"Mouse Controls:\n" +
                                   $"• Right-click: Context Menu\n" +
                                   $"• Hover: Show Title Bar (when not click-through)\n\n" +
                                   $"Color Schemes:\n" +
                                   $"1. Bass Heavy (Red→Yellow)\n" +
                                   $"2. Neon (Purple→Cyan)\n" +
                                   $"3. Fire (Dark Red→White)\n" +
                                   $"4. Ocean (Dark Blue→Cyan)\n\n" +
                                   $"💡 TIPS:\n" +
                                   $"• Press 'T' to toggle click-through mode\n" +
                                   $"• Right-click anywhere for quick access menu\n" +
                                   $"• Hover mouse to show window controls\n" +
                                   $"• Press 'H' to manually show/hide title bar";
            }
        }

        private void CycleColorScheme()
        {
            currentColorScheme = (currentColorScheme + 1) % colorSchemes.Length;
            if (showSettings) UpdateSettingsText();
        }

        private void AdjustSensitivity(double multiplier)
        {
            sensitivity = Math.Max(0.1, Math.Min(5.0, sensitivity * multiplier));
            if (showSettings) UpdateSettingsText();
        }

        private void AdjustBassBoost(double multiplier)
        {
            bassBoost = Math.Max(1.0, Math.Min(15.0, bassBoost * multiplier));
            if (showSettings) UpdateSettingsText();
        }

        private void SetBarCount(int newCount)
        {
            BAR_COUNT = newCount;
            InitializeArrays();
            CreateVisualizerBars();
            if (showSettings) UpdateSettingsText();
        }

        private void ToggleClickThrough()
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            if (hwnd != IntPtr.Zero)
            {
                int extendedStyle = GetWindowLong(hwnd, GWL_EXSTYLE);

                if (isClickThrough)
                {
                    SetWindowLong(hwnd, GWL_EXSTYLE, extendedStyle & ~WS_EX_TRANSPARENT);
                    this.ShowInTaskbar = true;
                }
                else
                {
                    SetWindowLong(hwnd, GWL_EXSTYLE, extendedStyle | WS_EX_TRANSPARENT);
                    this.ShowInTaskbar = false;
                }

                isClickThrough = !isClickThrough;
                if (showSettings) UpdateSettingsText();

                // Debug output
                System.Diagnostics.Debug.WriteLine($"Click-through toggled: {isClickThrough}");

                // Hide title bar when enabling click-through
                if (isClickThrough)
                {
                    TitleBar.Visibility = Visibility.Collapsed;
                    showTitleBar = false;
                }
            }
        }

        private void RestartAudioCapture()
        {
            StopAudioCapture();
            System.Threading.Thread.Sleep(200);
            StartAudioCapture();
        }

        private void OnWindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StopAudioCapture();
            renderTimer?.Stop();
            titleBarTimer?.Stop();
        }

        private void StopAudioCapture()
        {
            try
            {
                audioCapture?.StopRecording();
                audioCapture?.Dispose();
                audioCapture = null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error stopping audio capture: {ex.Message}");
            }
        }
        #endregion
    }
}