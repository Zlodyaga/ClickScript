using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CapsLockClicker
{
    public partial class MainForm : Form
    {
        private Random random = new Random();
        private System.Windows.Forms.Timer timer;
        private bool isPaused = false;
        private IntPtr _keyboardHookID = IntPtr.Zero;
        private IntPtr _mouseHookID = IntPtr.Zero;
        private LowLevelKeyboardProc _keyboardProc;
        private LowLevelMouseProc _mouseProc;

        private int timeForPause = convertNumberToSeconds(5);
        private int startPeriodForRandom = convertNumberToMinutes(3);
        private int endPeriodForRandom = convertNumberToMinutes(6);

        public MainForm()
        {
            InitializeComponent();

            // Initialize of hooks
            _keyboardProc = HookCallback;
            _mouseProc = MouseHookCallback;
            _keyboardHookID = SetHook(_keyboardProc);
            _mouseHookID = SetMouseHook(_mouseProc);

            // Set-up of timer
            timer = new System.Windows.Forms.Timer();
            timer.Tick += Timer_Tick;
            SetNextInterval();
            timer.Start();

            LogDebug("App started.");
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (isPaused) return; // If pause active - don't click

            ClickCapsLock();
            Task.Delay(1000).Wait();
            ClickCapsLock();

            LogDebug("Caps Lock was clicked twice.");

            SetNextInterval(); // Set next interval for click
        }

        private void SetNextInterval()
        {
            int interval = random.Next(startPeriodForRandom, endPeriodForRandom);
            timer.Interval = interval;
            LogDebug($"Next click will be in {interval / (convertNumberToSeconds(60))} minutes ({interval / 1000} seconds)");
        }

        private void ClickCapsLock()
        {
            keybd_event((byte)Keys.CapsLock, 0, 0, 0);
            keybd_event((byte)Keys.CapsLock, 0, 2, 0);
        }

        // Hook set-up
        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private IntPtr SetMouseHook(LowLevelMouseProc proc)
        {
            using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        // Callback for keyboard
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                PauseAndResetTimer();
            }
            return CallNextHookEx(_keyboardHookID, nCode, wParam, lParam);
        }

        // Callback for mouse (ONLY clicks!)
        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (wParam == (IntPtr)0x201 || wParam == (IntPtr)0x204)) // 0x201 - left button click or 0x204 - right button click
            {
                PauseAndResetTimer();
            }
            return CallNextHookEx(_mouseHookID, nCode, wParam, lParam);
        }
        
        private async void PauseAndResetTimer()
        {
            if (!isPaused)
            {
                isPaused = true;
                timer.Stop();
                LogDebug($"Activity of user was detected. Timer stopped for {timeForPause / 1000} seconds");

                await Task.Delay(timeForPause); // Async pause

                LogDebug("Timer repaused.");
                isPaused = false;
                timer.Start();
                SetNextInterval();
            }
        }


        // Method for logging in Debug
        private void LogDebug(string message) => Debug.WriteLine($"{DateTime.Now:HH:mm:ss} - {message}");

        private static int convertNumberToSeconds(int number) => number * 1000;

        private static int convertNumberToMinutes(int number) => convertNumberToSeconds(number) * 60;

        // Imports WinAPI
        [DllImport("user32.dll", SetLastError = true)]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        private const int WH_KEYBOARD_LL = 13;
        private const int WH_MOUSE_LL = 14;
    }
}
