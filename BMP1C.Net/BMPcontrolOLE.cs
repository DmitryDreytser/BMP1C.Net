using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Collections.Generic;
using BMP1C.Net._1CInterfaces;
using stdole;

namespace BMP1C.Net
{

    public class EnumOLEVERB : IEnumOLEVERB
    {
        void IEnumOLEVERB.Clone(out IEnumOLEVERB ppenum)
        {
            throw new NotImplementedException();
        }

        HRESULT IEnumOLEVERB.Next(int celt, tagOLEVERB rgelt, int[] pceltFetched)
        {
            throw new NotImplementedException();
        }

        void IEnumOLEVERB.Reset()
        {
            throw new NotImplementedException();
        }

        HRESULT IEnumOLEVERB.Skip(int celt)
        {
            throw new NotImplementedException();
        }

        public EnumOLEVERB()
        {

        }
    }

    [ProgId("BMP1C.Net")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComDefaultInterface(typeof(IBmpControl))]
    [ComSourceInterfaces(typeof(IBmpControlEvents))]
    [Guid("1051E56D-660D-4D05-A125-FC51A38D3D45")]
    [ComVisible(true)]
    public partial class BmPcontrol : IBmpControl, IOleObject, IOleControl, ICustomQueryInterface, IInitDone,
        IPersistStorage, IPersistStreamInit, IDisposable /*,IOleClientSite, IPersistMemory , IOleInPlaceSite, IOleLink*/
    {


        public EnumOLEVERB OleVerb = new EnumOLEVERB();
        private IOleAdviseHolder _AdviseHolder;

        private string BmpFile;
        private bool bSaved = false;

        //private DVASPECT dwDrawAspect;
        private uint dwDrawAspect;
        private string FileName;

        private Bitmap Image;

        private bool isDebuggerPresent = Debugger.IsAttached;

        private int m_Version = 2000;
        private bool NoDraw = false;

        private IOleClientSite pClientSite;
        private IMoniker pmk;

        private IStorage pstg;
        private tagSIZEL SiezeL;
        //private IntPtr SiezeL;

        public event OnDrawEventHandler OnDraw;

        string IBmpControl.BmpFile
        {
            get { return BmpFile; }

            set
            {
                BmpFile = value;
                //Exports.ImageRemoveRef(FileName);
                FileName = Path.GetFileName(BmpFile);
                if (File.Exists(BmpFile) || Exports.ImageCash.ContainsKey(FileName))
                {
                    if (Exports.ImageCash.ContainsKey(FileName))
                    {
                        Image = Exports.ImageCash[FileName];
                        if (isDebuggerPresent) Debug.Print("{0}: loaded {1} from Cache", GetType().GetCustomAttribute<ProgIdAttribute>().Value, FileName);
                    }
                    else
                    {
                        Image = new Bitmap(BmpFile);
                        Image.MakeTransparent(Color.White);
                       // Exports.ImageCash.Add(FileName, Image);
                        if (isDebuggerPresent) Debug.Print("{0}: loaded {1} from Disk", GetType().GetCustomAttribute<ProgIdAttribute>().Value, BmpFile);
                    }

                    //Exports.ImageAddRef(FileName);
                    Size = Image.Size;
                }
            }
        }

        int IBmpControl.Function { get; set; }
        int IBmpControl.GrMode { get; set; }

        int IBmpControl.NoDraw
        {
            get { return NoDraw ? 1 : 0; }

            set { NoDraw = value == 1; }
        }

        int IBmpControl.DstWidth { get; set; }

        int IBmpControl.DstHeight { get; set; }

        int IBmpControl.GroundClip { get; set; }

        int IBmpControl.SrcPointX { get; set; }

        int IBmpControl.SrcPointY { get; set; }

        int IBmpControl.SrcWidth { get; set; }

        int IBmpControl.SrcHeight { get; set; }

        int IBmpControl.GroundGor { get; set; }

        int IBmpControl.GroundVer { get; set; }

        int IBmpControl.DstDeltaPointX { get; set; }

        int IBmpControl.DstDeltaPointY { get; set; }

        [ComRegisterFunction]
        public static void RegisterClass(string key)
        {
            using (RegistryKey k = Registry.ClassesRoot.OpenSubKey(TrimClassesRoot(key), true))
            {
                RegistryKey ctrl = k.CreateSubKey("Control");
                ctrl.Close();

                RegistryKey inprocServer32 = k.OpenSubKey("InprocServer32", true);
                inprocServer32.SetValue("CodeBase", Assembly.GetExecutingAssembly().CodeBase);
                inprocServer32.Close();
                // Finally close the main	key
                k.Close();
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterClass(string key)
        {
            using (RegistryKey k = Registry.ClassesRoot.OpenSubKey(TrimClassesRoot(key), true))
            {
                k.DeleteSubKey("Control", false);
                // Next	open up	InprocServer32
                //RegistryKey	inprocServer32 = 
                k.OpenSubKey("InprocServer32", true);

                // And delete the CodeBase key,	again not throwing if missing
                k.DeleteSubKey("CodeBase", false);

                // Finally close the main key
                k.Close();
            }
        }

        private static string TrimClassesRoot(string key)
        {
            return key.Replace(@"HKEY_CLASSES_ROOT\", "");
        }

        void IOleObject.Advise(IAdviseSink pAdvSink, out uint pdwConnection)
        {
            if (_AdviseHolder == null) Natives.CreateOleAdviseHolder(out _AdviseHolder);
            _AdviseHolder.Advise(pAdvSink, out pdwConnection);
            if (isDebuggerPresent) Debug.Print("{0}: Advise", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }



        private void BMPcontrol_Paint(object sender, PaintEventArgs e)
        {
            if (isDebuggerPresent) Debug.Print("{0}: Draw", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
            if (Image == null || NoDraw)
                return;

            Random Rnd = new Random();
            Graphics gr = e.Graphics;
            gr.CompositingMode = CompositingMode.SourceOver;
            gr.SmoothingMode = SmoothingMode.AntiAlias;

            float cx = Width / 2.0f * (float)Rnd.NextDouble();
            float cy = Height / 2.0f * (float)Rnd.NextDouble();

            gr.TranslateTransform(cx, cy);
            float angle = (float)Rnd.Next(1, 3);
            gr.RotateTransform(angle);
            gr.TranslateTransform(-cx, -cy);
            gr.DrawImage(Image, e.ClipRectangle);
            gr.Flush();
            OnDraw?.Invoke(gr, e.ClipRectangle);
        }

        void IOleObject.Close(uint dwSaveOption)
        {

            if (isDebuggerPresent) Debug.Print("{0}: Close", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
            Exports.ImageRemoveRef(FileName);
            Image = null;

            if (pmk != null)
            {
                Marshal.ReleaseComObject(pmk);
                Marshal.FinalReleaseComObject(pmk);
                pmk = null;
            }

            if (pClientSite != null)
            {
                Marshal.ReleaseComObject(pClientSite);
                Marshal.FinalReleaseComObject(pClientSite);
                pClientSite = null;
            }

            if (pstg != null)
            {
                Marshal.ReleaseComObject(pstg);
                Marshal.FinalReleaseComObject(pstg);
                pstg = null;
            }

            Dispose(true);
            //GC.Collect();
            //GC.Collect();
            //GC.SuppressFinalize(this);
            //GC.WaitForPendingFinalizers();
            
        }

        void IInitDone.Done()
        {
            Marshal.ReleaseComObject(Exports.V7Object);
            Marshal.FinalReleaseComObject(Exports.V7Object);
            Exports.V7Object = null;
            Marshal.ReleaseComObject(Exports.asyncEvent);
            Marshal.FinalReleaseComObject(Exports.asyncEvent);
            Exports.asyncEvent = null;
            Marshal.ReleaseComObject(Exports.statusLine);
            Marshal.FinalReleaseComObject(Exports.statusLine);
            Exports.statusLine = null;
            Marshal.ReleaseComObject(Exports.connection);
            Marshal.FinalReleaseComObject(Exports.connection);
            Exports.connection = null;
            //GC.Collect();
            //GC.Collect();
        }

        void IOleObject.DoVerb(OleVerbs iVerb, IntPtr lpmsg, IOleClientSite pActiveSite, int lindex, IntPtr hwndParent,
            ref RECT lprcPosRect)
        {
            if (isDebuggerPresent) Debug.Print("{0}: DoVerb", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IOleObject.EnumAdvise(out IEnumStatData ppenumAdvise)
        {
            if (_AdviseHolder == null) Natives.CreateOleAdviseHolder(out _AdviseHolder);
            _AdviseHolder.EnumAdvise(out ppenumAdvise);
            if (isDebuggerPresent) Debug.Print("{0}: EnumAdvise", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

      
        //void IOleObject.EnumVerbs(out IEnumOLEVERB ppEnumOleVerb)
        void IOleObject.EnumVerbs(out IEnumOLEVERB ppEnumOleVerb)
        {
            ppEnumOleVerb = OleVerb;
            if (isDebuggerPresent) Debug.Print("{0}: EnumVerbs", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IOleControl.FreezeEvents(bool bFreeze)
        {
            if (isDebuggerPresent) Debug.Print("{0}: FreezeEvents", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IPersistStorage.GetClassID(out Guid pClassID)
        {
            ((IPersistStreamInit)this).GetClassID(out pClassID);
        }

        int IPersist.GetClassID(out Guid pClassID)
        {
            pClassID = GetGuid();
            return 0;
        }

        void IOleObject.GetClientSite(out IOleClientSite ppClientSite)
        {
            ppClientSite = pClientSite;
            if (isDebuggerPresent) Debug.Print("{0}: GetClientSite", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IOleObject.GetClipboardData(uint dwReserved, out IDataObject ppDataObject)
        {
            ppDataObject = null;
            if (isDebuggerPresent) Debug.Print("{0}: GetClipboardData", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IOleControl.GetControlInfo(ref CONTROLINFO pCI)
        {
            if (isDebuggerPresent) Debug.Print("{0}: GetControlInfo", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IOleObject.GetExtent(uint dwDrawAspect, tagSIZEL psizel)
        {
            //psizel = this.SiezeL;
            if (isDebuggerPresent) Debug.Print("{0}: GetExtent", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        private Guid GetGuid()
        {
            return new Guid(GetType().GetCustomAttribute<GuidAttribute>().Value);
        }

        void IInitDone.GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info)
        {
            info[0] = this.m_Version;
            info[1] = "BMP1C.Net";
        }

        CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;

            if (iid == typeof(IPersistStorage).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IPersistStorage), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }

            if (iid == typeof(IOleControl).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IOleControl), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }

            if (iid == typeof(IPersistStreamInit).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IPersistStreamInit),
                    CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }


            if (iid == typeof(IOleObject).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IOleObject), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }

            if (iid == typeof(IInitDone).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IInitDone),
                    CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }

            //Type[] ints = this.GetType().GetInterfaces();
            //foreach (Type currentInterface in ints)
            //{
            //    if (currentInterface.GUID == iid)
            //    {
            //        ppv = Marshal.GetComInterfaceForObject(this, currentInterface, CustomQueryInterfaceMode.Ignore);
            //        if (isDebuggerPresent) Debug.Print("{0}: Interface {1} is handled", GetType().GetCustomAttribute<ProgIdAttribute>().Value, currentInterface.Name);
            //        return CustomQueryInterfaceResult.Handled;
            //    }
            //}

            if (isDebuggerPresent) Debug.Print("{0}: Interface {1} is not handled", GetType().GetCustomAttribute<ProgIdAttribute>().Value, iid);
            return CustomQueryInterfaceResult.NotHandled;
        }

        void IOleObject.GetMiscStatus(DVASPECT dwAspect, out OLEMISC pdwStatus)
        {
            pdwStatus = OLEMISC.OLEMISC_RECOMPOSEONRESIZE | OLEMISC.OLEMISC_RENDERINGISDEVICEINDEPENDENT;

            if (isDebuggerPresent) Debug.Print("{0}: GetMiscStatus", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IOleObject.GetMoniker(uint dwAssign, uint dwWhichMoniker, out IMoniker ppmk)
        {
            ppmk = pmk;
            if (isDebuggerPresent) Debug.Print("{0}: GetMoniker", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        int IPersistStreamInit.GetSizeMax(out long pCbSize)
        {
            pCbSize = 10 * 1024 * 1024;
            return 0;
        }

        void IOleObject.GetUserClassID(out Guid pClsid)
        {
            ((IPersistStreamInit)this).GetClassID(out pClsid);
        }

        void IOleObject.GetUserType(uint dwFormOfType, out string pszUserType)
        {
            pszUserType = null;
            if (isDebuggerPresent) Debug.Print("{0}: GetUserType", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        uint IPersistStorage.HandsOffStorage()
        {
            return 0;
        }

        void IInitDone.Init([MarshalAs(UnmanagedType.IDispatch)]dynamic connection)
        {
            Exports.connection = connection;
            Exports.V7Object = connection.AppDispatch;
            Exports.asyncEvent = (IAsyncEvent)connection;
            Exports.statusLine = (IStatusLine)connection;
        }

        void IOleObject.InitFromData(IDataObject pDataObject, bool fCreation, uint dwReserved)
        {
            if (isDebuggerPresent) Debug.Print("{0}: InitFromData", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IPersistStorage.InitNew(IStorage pstg)
        {
            this.pstg = pstg;
            BmpFile = string.Empty;
            Image = null;
            if (isDebuggerPresent) Debug.Print("{0}: InitNew", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        int IPersistStreamInit.InitNew()
        {
            BmpFile = string.Empty;
            Image = null;
            return 0;
        }

        bool IPersistStorage.IsDirty()
        {
            if (((IPersistStreamInit)this).IsDirty()) ((IPersistStorage)this).Save(pstg, true);

            return !bSaved;
        }

        bool IPersistStreamInit.IsDirty()
        {
            return !bSaved;
        }

        void IOleObject.IsUpToDate()
        {
            if (isDebuggerPresent) Debug.Print("{0}: IsUpToDate", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        uint IPersistStorage.Load(IStorage pstg)
        {
            IEnumSTATSTG stg = null;
            var res = pstg.EnumElements(0, IntPtr.Zero, 0, out stg);
            if (stg != null)
            {
                STATSTG[] regelt = { new STATSTG() };

                uint fetched = 0;

                while (stg.Next(1, regelt, out fetched) == 0)
                    if (regelt[0].type == STGTY.STGTY_STREAM)
                    {
                        BmpFile = regelt[0].pwcsName;
                        FileName = regelt[0].pwcsName;
                        if (Exports.ImageCash.ContainsKey(FileName))
                        {
                            Image = Exports.ImageCash[FileName];
                            if (isDebuggerPresent) Debug.Print("{0}: Loaded from Cache", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
                        }
                        else
                        {
                            IStream PictureStream = null;
                            pstg.OpenStream(BmpFile, IntPtr.Zero, STGM.STGM_READWRITE | STGM.STGM_SHARE_EXCLUSIVE, 0,
                                out PictureStream);
                            byte[] imageData = new byte[regelt[0].cbSize];
                            PictureStream.Read(imageData, imageData.Length, IntPtr.Zero);

                            using (var ms = new MemoryStream(imageData))
                            {
                                Image = new Bitmap(ms);
                                Image.MakeTransparent(Color.LightGray);
                                Size = Image.Size;
                            }

                            Marshal.ReleaseComObject(PictureStream);
                            Marshal.FinalReleaseComObject(PictureStream);
                            PictureStream = null;

                            //GC.Collect();
                            //GC.Collect();
                            //GC.WaitForPendingFinalizers();
                            Exports.ImageCash.Add(FileName, Image);
                            if (isDebuggerPresent) Debug.Print("{0}: Loaded from Storage", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
                        }
                        Exports.ImageAddRef(FileName);

                    }

                return 0;
            }

            if (BmpFile == string.Empty)
                return 0x800401F1;
            return 1;
        }

        int IPersistStreamInit.Load(IStream pStm)
        {
            return 0;
        }

        void IOleControl.OnAmbientPropertyChange(int dispID)
        {
            if (isDebuggerPresent) Debug.Print("{0}: OnAmbientPropertyChange", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IOleControl.OnMnemonic(Message pMsg)
        {
            if (isDebuggerPresent) Debug.Print("{0}: OnMnemonic", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        uint IPersistStorage.Save(IStorage pStgSave, bool fSameAsLoad)
        {
            bSaved = true;

            if (Image == null)
                return (uint)HRESULT.S_OK;

            IStream pSaveStream = null;
            try
            {
                pStgSave.CreateStream(Path.GetFileName(BmpFile),
                    STGM.STGM_CREATE | STGM.STGM_WRITE | STGM.STGM_SHARE_EXCLUSIVE, 0, 0, out pSaveStream);

                if (pSaveStream != null)
                {
                    int len = 0;

                    IPicture pict = ImageUtilities.ConvertToIPicture(Image);
                    pict.SaveAsFile(Marshal.GetComInterfaceForObject(pSaveStream, typeof(IStream)), true, out len);
                    pSaveStream.Commit(1);

                    Marshal.ReleaseComObject(pict);
                    Marshal.FinalReleaseComObject(pict);
                    pict = null;
                    if (isDebuggerPresent) Debug.Print("{0}: Saved to Storage", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
                }

                if (pSaveStream != null)
                {
                    Marshal.ReleaseComObject(pSaveStream);
                    Marshal.FinalReleaseComObject(pSaveStream);
                    pSaveStream = null;
                }
                

                pStgSave.Commit(STGC.OVERWRITE);
                Marshal.ReleaseComObject(pStgSave);
                Marshal.FinalReleaseComObject(pStgSave);
                //GC.Collect();
                //GC.Collect();
                //GC.WaitForPendingFinalizers();


            }
            catch (Exception ex)
            {
                return (uint) ex.HResult;
            }

            return (uint)HRESULT.S_OK;
        }

        int IPersistStreamInit.Save(IStream pStm, bool fClearDirty)
        {
            return 0;
        }

        uint IPersistStorage.SaveCompleted(IStorage pStgNew)
        {
            return 0;
        }

        void IOleObject.SetClientSite(IOleClientSite pClientSite)
        {
            this.pClientSite = pClientSite;
            if (isDebuggerPresent) Debug.Print("{0}: SetClientSite", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IOleObject.SetColorScheme(object pLogpal)
        {
            if (isDebuggerPresent) Debug.Print("{0}: SetColorScheme", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }
        void IOleObject.SetExtent(uint dwDrawAspect, tagSIZEL psizel)
        {
            this.dwDrawAspect = dwDrawAspect;
            this.SiezeL = psizel;

            if (isDebuggerPresent) Debug.Print("{0}: SetExtent", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IOleObject.SetHostNames(string szContainerApp, string szContainerObj)
        {
            if (isDebuggerPresent) Debug.Print("{0}: SetHostNames", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IOleObject.SetMoniker(uint dwWhichMoniker, IMoniker pmk)
        {
            this.pmk = pmk;
            if (isDebuggerPresent) Debug.Print("{0}: SetMoniker", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IOleObject.Unadvise(uint dwConnection)
        {
            if (_AdviseHolder == null) Natives.CreateOleAdviseHolder(out _AdviseHolder);
            _AdviseHolder.Unadvise(dwConnection);
            if (isDebuggerPresent) Debug.Print("{0}: Unadvise", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

        void IOleObject.Update()
        {
            if (isDebuggerPresent) Debug.Print("{0}: Update", GetType().GetCustomAttribute<ProgIdAttribute>().Value);
        }

    }

    static class Exports
    {
        public static IAsyncEvent asyncEvent;
        public static IStatusLine statusLine;
        public static dynamic connection;
        public static dynamic V7Object;
        public static Dictionary<string, Bitmap> ImageCash = new Dictionary<string, Bitmap>();
        public static Dictionary<string, int> ImageRefs = new Dictionary<string, int>();

        public static void ImageAddRef(string FileName)
        {

            if (Exports.ImageRefs.ContainsKey(FileName))
                Exports.ImageRefs[FileName]++;
            else
                Exports.ImageRefs.Add(FileName, 1);
        }

        public static void ImageRemoveRef(string FileName)
        {
            if (FileName != null)
                if (Exports.ImageRefs.ContainsKey(FileName))
                {
                    Exports.ImageRefs[FileName]--;
                    if (Exports.ImageRefs[FileName] == 0)
                    {
                        Exports.ImageRefs.Remove(FileName);
                        Exports.ImageCash[FileName].Dispose();
                        Exports.ImageCash.Remove(FileName);
                        Debug.Print("BMP1C.Net: {0} removed from cache", FileName);
                    }
                }
        }
    }
}


