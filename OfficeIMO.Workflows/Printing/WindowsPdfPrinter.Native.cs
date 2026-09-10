using System.Runtime.InteropServices;
using System.Text;

namespace OfficeIMO.Workflows;

internal static partial class WindowsPdfPrinter {
    private static class Native {
        [DllImport("winspool.drv", EntryPoint = "DeviceCapabilitiesW", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int DeviceCapabilities(string device, string? port, ushort capability, IntPtr output, IntPtr mode);
        [StructLayout(LayoutKind.Sequential)]
        internal struct PrinterInfo2 {
            internal IntPtr ServerName, PrinterName, ShareName, PortName, DriverName, Comment, Location, DevMode;
            internal IntPtr SepFile, PrintProcessor, DataType, Parameters, SecurityDescriptor;
            internal uint Attributes, Priority, DefaultPriority, StartTime, UntilTime, Status, Jobs, AveragePpm;
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        internal struct DocInfo {
            internal int Size;
            [MarshalAs(UnmanagedType.LPWStr)] internal string DocumentName;
            [MarshalAs(UnmanagedType.LPWStr)] internal string? Output;
            [MarshalAs(UnmanagedType.LPWStr)] internal string? DataType;
            internal int Type;
        }
        [StructLayout(LayoutKind.Sequential)]
        internal struct BitmapInfo {
            internal uint Size;
            internal int Width, Height;
            internal ushort Planes, BitCount;
            internal uint Compression, ImageSize;
            internal int XPixelsPerMeter, YPixelsPerMeter;
            internal uint ColorsUsed, ColorsImportant;
        }
        [DllImport("winspool.drv", EntryPoint = "EnumPrintersW", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool EnumPrinters(uint flags, string? name, uint level, IntPtr buffer, uint size, out uint needed, out uint returned);
        [DllImport("winspool.drv", EntryPoint = "GetDefaultPrinterW", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetDefaultPrinter(StringBuilder? name, ref uint characters);
        [DllImport("winspool.drv", EntryPoint = "OpenPrinterW", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool OpenPrinter(string name, out IntPtr handle, IntPtr defaults);
        [DllImport("winspool.drv", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ClosePrinter(IntPtr handle);
        [DllImport("winspool.drv", EntryPoint = "DocumentPropertiesW", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int DocumentProperties(IntPtr window, IntPtr printer, string name, IntPtr output, IntPtr input, uint mode);
        [DllImport("gdi32.dll", EntryPoint = "CreateDCW", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr CreateDC(string driver, string device, string? output, IntPtr mode);
        [DllImport("gdi32.dll", EntryPoint = "ResetDCW", SetLastError = true)] internal static extern IntPtr ResetDC(IntPtr dc, IntPtr mode);
        [DllImport("gdi32.dll", EntryPoint = "StartDocW", CharSet = CharSet.Unicode, SetLastError = true)] internal static extern int StartDoc(IntPtr dc, ref DocInfo info);
        [DllImport("gdi32.dll", SetLastError = true)] internal static extern int StartPage(IntPtr dc);
        [DllImport("gdi32.dll", SetLastError = true)] internal static extern int EndPage(IntPtr dc);
        [DllImport("gdi32.dll", SetLastError = true)] internal static extern int EndDoc(IntPtr dc);
        [DllImport("gdi32.dll", SetLastError = true)] internal static extern int AbortDoc(IntPtr dc);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal delegate bool AbortProcedure(IntPtr dc, int error);
        [DllImport("gdi32.dll", SetLastError = true)] internal static extern int SetAbortProc(IntPtr dc, AbortProcedure callback);
        [DllImport("gdi32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool DeleteDC(IntPtr dc);
        [DllImport("gdi32.dll")] internal static extern int GetDeviceCaps(IntPtr dc, int index);
        [DllImport("gdi32.dll", SetLastError = true)]
        internal static extern int StretchDIBits(IntPtr dc, int x, int y, int width, int height, int sourceX, int sourceY,
            int sourceWidth, int sourceHeight, byte[] bits, ref BitmapInfo info, uint usage, uint operation);
    }
}
