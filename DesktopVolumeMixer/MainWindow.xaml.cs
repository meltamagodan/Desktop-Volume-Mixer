using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using NAudio.CoreAudioApi;
using System;
using System.Runtime.InteropServices;
using System.Text;
using WinRT.Interop;
using AppWindow = Microsoft.UI.Windowing.AppWindow;


namespace DesktopVolumeMixer
{
    public sealed partial class MainWindow : Window
    {
        // Fields
        private readonly MMDeviceEnumerator _deviceEnumerator;
        private readonly MMDevice _defaultDevice;
        private readonly AudioSessionManager _sessionManager;
        private readonly nint _hWnd;
        private readonly OverlappedPresenter _presenter;
        private readonly DispatcherTimer _refreshTimer;
        private readonly AppWindow _appWindow;
        private SessionCollection? sessions;

        // P/Invoke methods
        [DllImport("user32.dll")]
        static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        static extern void SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOOLWINDOW = 0x00000080;

        public MainWindow()
        {
            InitializeComponent();

            // Initialize window and audio components
            _hWnd = WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(_hWnd);
            _appWindow = AppWindow.GetFromWindowId(windowId);
            SetWindowStyle();
            _appWindow.Resize(new Windows.Graphics.SizeInt32 { Width = 270, Height = 500 });

            _presenter = (OverlappedPresenter)_appWindow.Presenter;
            ConfigureWindowPresenter();

            ExtendsContentIntoTitleBar = true;

            // Initialize audio devices
            _deviceEnumerator = new MMDeviceEnumerator();
            _defaultDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            _sessionManager = _defaultDevice.AudioSessionManager;

            // Load audio sessions
            LoadAudioSessions();

            // Set up the refresh timer
            _refreshTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
            _refreshTimer.Tick += RefreshTimer_Tick;
            _refreshTimer.Start();
        }

        private void SetWindowStyle()
        {
            int exStyle = GetWindowLong(_hWnd, GWL_EXSTYLE);
            SetWindowLong(_hWnd, GWL_EXSTYLE, exStyle | WS_EX_TOOLWINDOW);
        }

        private void ConfigureWindowPresenter()
        {
            _presenter.IsMaximizable = false;
            _presenter.IsMinimizable = false;
            _presenter.IsResizable = false;
        }

        private void RefreshTimer_Tick(object? sender, object e)
        {
            IntPtr foregroundHwnd = GetForegroundWindow();
            StringBuilder title = new StringBuilder(256);
            GetWindowText(foregroundHwnd, title, title.Capacity);

            // Check if the foreground window is Program Manager(Desktop) and prevent Show Desktop from hinding the Window
            if (title.ToString() == "Program Manager" || title.ToString() == "Desktop Volumater")
            {
                _presenter.IsAlwaysOnTop = true;
            } else
            {
                _presenter.IsAlwaysOnTop = false;
                SetWindowPos(_hWnd, HWND_BOTTOM, 0, 0, 0 ,0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
            }

            // Refresh audio sessions
            LoadAudioSessions();
        }


        private void LoadAudioSessions()
        {
            try
            {
                _defaultDevice.AudioSessionManager.RefreshSessions();
                var refresh = _defaultDevice.AudioSessionManager.Sessions;
                if (sessions != null && refresh.Count == sessions.Count)
                {
                    return;
                }

                SessionList.Children.Clear();
                sessions = _defaultDevice.AudioSessionManager.Sessions;

                for (int i = 0; i < sessions.Count; i++)
                {
                    var session = sessions[i];
                    var volume = session.SimpleAudioVolume;
                    string label;

                    // Determine label
                    if (session.IsSystemSoundsSession)
                    {
                        label = "System Sounds";
                    }
                    else if (!string.IsNullOrWhiteSpace(session.DisplayName))
                    {
                        label = session.DisplayName;
                    }
                    else
                    {
                        try
                        {
                            var process = System.Diagnostics.Process.GetProcessById((int)session.GetProcessID);
                            label = process.ProcessName;
                        }
                        catch
                        {
                            label = $"PID {session.GetProcessID}";
                        }
                    }

                    // Create UI elements
                    var nameText = new TextBlock
                    {
                        Text = $"{label} ({(int)(volume.Volume * 100)}%)",
                        Margin = new Thickness(0, 0, 0, 5),
                        FontSize = 16
                    };

                    var slider = new Slider
                    {
                        Minimum = 0,
                        Maximum = 100,
                        Value = volume.Volume * 100,
                        Width = 200,
                        HorizontalAlignment = HorizontalAlignment.Left
                    };

                    slider.ValueChanged += (s, e) =>
                    {
                        try
                        {
                            volume.Volume = (float)(e.NewValue / 100);
                            nameText.Text = $"{label} ({(int)e.NewValue}%)";
                        }
                        catch (Exception ex)
                        {
                            nameText.Text = $"{label} (Error)";
                            System.Diagnostics.Debug.WriteLine($"Volume set error: {ex.Message}");
                        }
                    };

                    var container = new StackPanel();

                    container.Children.Add(nameText);
                    container.Children.Add(slider);
                    SessionList.Children.Add(container);
                }

                if (SessionList.Children.Count == 0)
                {
                    SessionList.Children.Add(new TextBlock
                    {
                        Text = "No active audio sessions.",
                        FontSize = 16,
                        Foreground = new SolidColorBrush(Colors.Gray)
                    });
                }
            }
            catch (Exception ex)
            {
                SessionList.Children.Clear();
                SessionList.Children.Add(new TextBlock
                {
                    Text = $"Error loading sessions: {ex.Message}",
                    FontSize = 16,
                    Foreground = new SolidColorBrush(Colors.Red)
                });
            }
        }
    }
}