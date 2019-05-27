using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Linq;

namespace BMP1C.Net
{
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct StgMedium
    {
        public uint tymed;
        public IntPtr unionmember;
        [MarshalAs(UnmanagedType.IUnknown)]
        public object pUnkForRelease;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct FormatEtc
    {
        public ushort cfFormat;
        public IntPtr ptd;
        public uint dwAspect;
        public int lindex;
        public uint tymed;
    }

    [Guid("0000010E-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IDataObject
    {
        [PreserveSig]
        int GetData(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pformatetcIn,
            [MarshalAs(UnmanagedType.LPArray)] [Out] StgMedium[] pRemoteMedium);
        [PreserveSig]
        int GetDataHere(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc,
            [MarshalAs(UnmanagedType.LPArray)] [In] [Out] StgMedium[] pRemoteMedium);
        [PreserveSig]
        int QueryGetData(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc);
        [PreserveSig]
        int GetCanonicalFormatEtc(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pformatectIn,
            [MarshalAs(UnmanagedType.LPArray)] [Out] FormatEtc[] pformatetcOut);
        [PreserveSig]
        int SetData(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc,
            [MarshalAs(UnmanagedType.LPArray)] [In] StgMedium[] pmedium,
            [In] int fRelease);
        [PreserveSig]
        int EnumFormatEtc(
            [In] uint dwDirection,
            [MarshalAs(UnmanagedType.Interface)] out IEnumFormatEtc ppenumFormatEtc);
        [PreserveSig]
        int DAdvise(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc,
            [In] uint ADVF,
            [MarshalAs(UnmanagedType.Interface)] [In] IAdviseSink pAdvSink,
            out uint pdwConnection);
        [PreserveSig]
        int DUnadvise(
            [In] uint dwConnection);
        [PreserveSig]
        int EnumDAdvise(
            [MarshalAs(UnmanagedType.Interface)] out IEnumStatData ppenumAdvise);
    }
}
