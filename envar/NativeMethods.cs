using System;
using System.Runtime.InteropServices;

namespace envar
{
    internal class NativeMethods
    {
        public static IntPtr HWND_BROADCAST = new IntPtr(0xffff);
        public const uint SMTO_ABORTIFHUNG = 0x0002;
        public const uint WM_SETTINGCHANGE = 0x001A;
        public const ulong ERROR_TIMEOUT = 1460L;

        [DllImport("User32.dll", EntryPoint = "SendMessageTimeout", CallingConvention = CallingConvention.Winapi, CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr SendMessageTimeout(
            IntPtr hWnd,
            uint Msg,
            UIntPtr wParam,
            IntPtr lParam,
            uint fuFlags,
            uint uTimeout,
            out UIntPtr lpdwResult);
    }
}
