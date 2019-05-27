using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace BMP1C.Net
{
    internal class Natives
    {

        [DllImport("ole32.dll")]
        public static extern int OleInitialize(IntPtr pvReserved);

        [DllImport("ole32.dll")]
        public static extern void OleUninitialize();

        [DllImport("ole32.dll")]
        public static extern void CoFreeUnusedLibrariesEx(int delay, int dwReserved);

        [DllImport("ole32.dll")]
        public static extern void CoFreeLibrary(IntPtr hInst);

        [DllImport("ole32.dll")]
        public static extern HRESULT DllCanUnloadNow();

        [DllImport("ole32.dll")]
        public static extern int CreateOleAdviseHolder
            ([MarshalAs(UnmanagedType.Interface)] out IOleAdviseHolder ppOAHolder);
        [DllImport("ole32.dll")]
        public static extern int StgCreateDocfile([MarshalAs(UnmanagedType.LPWStr)]
                string pwcsName, int grfMode, uint reserved,
            [MarshalAs(UnmanagedType.Interface)] out IStorage ppstgOpen);
        [DllImport("ole32.dll")]
        public static extern int WriteClassStg(
            [MarshalAs(UnmanagedType.Interface)] IStorage pStg, ref Guid rclsid);
        [DllImport("ole32.dll")]
        public static extern int OleSetClipboard([MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IDataObject pDataObj);

        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateMetaFile(string lpszFile);
        [DllImport("gdi32.dll")]
        public static extern IntPtr CloseMetaFile(IntPtr hdc);
        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateDC(string lpszDriver, string lpszDevice,
           string lpszOutput, IntPtr lpInitData);
        [DllImport("gdiplus.dll")]
        public static extern uint GdipEmfToWmfBits(IntPtr _hEmf, uint _bufferSize,
            byte[] _buffer, int _mappingMode, EmfToWmfBitsFlags _flags);
        [DllImport("gdi32.dll")]
        public static extern IntPtr SetMetaFileBitsEx(uint _bufferSize,
            byte[] _buffer);

        public const int E_NOTIMPL = unchecked((int)0x80004001);
        public const int DV_E_FORMATETC = unchecked((int)0x80040064);

        [Flags]
        public enum EmfToWmfBitsFlags
        {
            EmfToWmfBitsFlagsDefault = 0x00000000,
            EmfToWmfBitsFlagsEmbedEmf = 0x00000001,
            EmfToWmfBitsFlagsIncludePlaceable = 0x00000002,
            EmfToWmfBitsFlagsNoXORClip = 0x00000004
        }

        public const int MM_ISOTROPIC = 7;
        public const int MM_ANISOTROPIC = 8;


        [DllImport("ole32.dll")]
        public static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);
    }
}
