using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.IO;

namespace BMP1C.Net
{

    [ComVisible(true)]
    [Guid("D783DC48-8FB0-4fe9-BDC2-0CEE3F5E8921")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IBmpControl
    {
        [DispId(4)]
        string BmpFile { get; set; }
        [DispId(5)]
        int GrMode { get; set; }
        [DispId(6)]
        int NoDraw { get; set; }
        [DispId(7)]
        int Function { get; set; }
        [DispId(8)]
        int DstWidth { get; set; }
        [DispId(9)]
        int DstHeight { get; set; }
        [DispId(10)]
        int GroundClip { get; set; }
        [DispId(11)]
        int SrcPointX { get; set; }
        [DispId(12)]
        int SrcPointY { get; set; }
        [DispId(13)]
        int SrcWidth { get; set; }
        [DispId(14)]
        int SrcHeight { get; set; }
        [DispId(15)]
        int GroundGor { get; set; }
        [DispId(16)]
        int GroundVer { get; set; }
        [DispId(17)]
        int DstDeltaPointX { get; set; }
        [DispId(18)]
        int DstDeltaPointY { get; set; }
    }

    [ComVisible(true)]
    [Guid("89948C6D-820E-4d19-94A5-587D57408391")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IBmpControlEvents
    {
        [DispId(2)]
        void OnDraw(Graphics gr, Rectangle rect);
    }

    public delegate void OnDrawEventHandler(Graphics gr, Rectangle rect);
}
