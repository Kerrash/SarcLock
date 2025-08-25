using System;
using System.Reflection;
using System.IO;
using System.Media;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using WindowsInput;

namespace SarcLock
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayApp());
        }
    }

    public class TrayApp : ApplicationContext
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        const int KEYEVENTF_EXTENDEDKEY = 0x1;
        const int KEYEVENTF_KEYUP = 0x2;
        const byte VK_CAPITAL = 0x14;
        private bool ignoreNextCaps = false;


        private void ToggleCapsLock()
        {
            ignoreNextCaps = true;
            keybd_event(VK_CAPITAL, 0x45, KEYEVENTF_EXTENDEDKEY, 0);
            keybd_event(VK_CAPITAL, 0x45, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
        }


        private NotifyIcon trayIcon;
        private GlobalKeyboardHook gHook;
        private InputSimulator sim = new InputSimulator();

        private int capsTapCount = 0;
        private DateTime lastTapTime = DateTime.MinValue;
        private bool flipMode = false;

        private Icon iconOn;
        private Icon iconOff;

        public TrayApp()
        {
            iconOn = LoadIcon("SarcLock.Icons.lock_on.ico");
            iconOff = LoadIcon("SarcLock.Icons.lock_off.ico");

            trayIcon = new NotifyIcon()
            {
                Icon = iconOff,
                Visible = true,
                Text = "SarcLock",
                ContextMenuStrip = new ContextMenuStrip()
            };
            trayIcon.ContextMenuStrip.Items.Add("Exit", null, Exit);

            gHook = new GlobalKeyboardHook();
            gHook.KeyDown += new KeyEventHandler(GHook_KeyDown);
            gHook.Hook();
        }

        private Icon LoadIcon(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            {
                return new Icon(stream);
            }
        }

        private void GHook_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.CapsLock)
            {
                if (ignoreNextCaps)
                {
                    ignoreNextCaps = false;
                    return; // Skip synthetic CapsLock events
                }

                var now = DateTime.Now;
                if ((now - lastTapTime).TotalMilliseconds < 500)
                    capsTapCount++;
                else
                    capsTapCount = 1;

                lastTapTime = now;

                if (capsTapCount == 3)
                {
                    flipMode = !flipMode;
                    capsTapCount = 0;
                    Debug.WriteLine("Flip mode: " + flipMode);

                    trayIcon.Icon = flipMode ? iconOn : iconOff;
                    trayIcon.Text = flipMode ? "SarcLock (ON)" : "SarcLock (OFF)";

                    if (flipMode)
                        SystemSounds.Asterisk.Play();   // Mode ON
                    else
                        SystemSounds.Hand.Play();       // Mode OFF
                }
            }
            else if (flipMode)
            {
                ToggleCapsLock();
            }
        }


        void Exit(object sender, EventArgs e)
        {
            gHook.Unhook();
            trayIcon.Visible = false;
            Application.Exit();
        }
    }

    public class GlobalKeyboardHook
    {
        [DllImport("user32.dll")] private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")] private static extern IntPtr GetModuleHandle(string lpModuleName);

        private IntPtr hookId = IntPtr.Zero;
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;

        public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private LowLevelKeyboardProc proc;

        public event KeyEventHandler KeyDown;

        public void Hook()
        {
            proc = HookCallback;
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                hookId = SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        public void Unhook()
        {
            UnhookWindowsHookEx(hookId);
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Keys key = (Keys)vkCode;
                KeyDown?.Invoke(this, new KeyEventArgs(key));
            }
            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }
    }
}
