using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using static Moxel.Moxel;

namespace Moxel
{
    public class EmbeddedObject : DataCell
    {

        public Rectangle AbsoluteImageArea
        {
            get
            {
                int left = 0, top = 0, right = 0, bottom = 0;

                left = (int)Math.Round(Parent.GetWidth(0, Picture.dwColumnStart) * 0.875 + Picture.dwOffsetLeft / 2.8, 0);
                right = (int)Math.Round(Parent.GetWidth(0, Picture.dwColumnEnd) * 0.875 + Picture.dwOffsetRight / 2.8, 0);

                top = (int)Math.Round(Parent.GetHeight(0, Picture.dwRowStart) / 3 + Picture.dwOffsetTop / 2.8, 0);
                bottom = (int)Math.Round(Parent.GetHeight(0, Picture.dwRowEnd) / 3 + Picture.dwOffsetBottom / 2.8, 0);

                return new Rectangle { X = Math.Min(left, right), Y = Math.Min(top, bottom), Width = Math.Abs(right - left), Height = Math.Abs(bottom - top) };
            }
        }

        public Rectangle ImageArea
        {
            get
            {
                int left = 0, top = 0, right = 0, bottom = 0;

                left = (int)Math.Round(Parent.GetWidth(0, Picture.dwColumnStart) * 0.875 + Picture.dwOffsetLeft / 2.8, 0);
                right = (int)Math.Round(Parent.GetWidth(0, Picture.dwColumnEnd) * 0.875 + Picture.dwOffsetRight / 2.8, 0);

                top = (int)Math.Round(Parent.GetHeight(0, Picture.dwRowStart) / 3 + Picture.dwOffsetTop / 2.8, 0);
                bottom = (int)Math.Round(Parent.GetHeight(0, Picture.dwRowEnd) / 3 + Picture.dwOffsetBottom / 2.8, 0);

                return new Rectangle { X = left, Y = top, Width = right - left, Height = bottom - top };
            }
        }
        public Picture Picture;
        public object pObject;
        public object OleObject;
        public string ProgID;
        public Guid ClsId;
        public byte[] OleObjectStorage;

        public EmbeddedObject()
        { }
        public EmbeddedObject(BinaryReader br, Moxel parent) : base(br, parent)
        {
            Picture = br.Read<Picture>();
            switch (Picture.dwType)
            {
                case ObjectType.Picture:
                    pObject = LoadPicture(br);
                    break;
                case ObjectType.Ole:
                    pObject = LoadOleObject(br);
                    break;
                default:
                    break;
            }
        }

        object LoadOleObject(BinaryReader br)
        {
            string classname = string.Empty;
            short wClassNameFlag = br.ReadInt16();
            if (wClassNameFlag == -1)
            {
                br.ReadInt16();
                short wClassNameLength = br.ReadInt16();
                classname = Encoding.GetEncoding(1251).GetString(br.ReadBytes(wClassNameLength));
            }

            int dwObjectType = br.ReadInt32();
            int dwItemNumber = br.ReadInt32();
            int dwAspect = br.ReadInt32();
            short wUseMoniker = br.ReadInt16();

            dwAspect = br.ReadInt32();

            int dwObjectSize = br.ReadInt32();
            OleObjectStorage = br.ReadBytes(dwObjectSize);

            OLE32.CoInitializeEx(IntPtr.Zero, OLE32.CoInit.ApartmentThreaded); //COINIT_APARTMENTTHREADED
            OLE32.ILockBytes LockBytes;
            OLE32.IStorage RootStorage;

            IntPtr hGlobal = Marshal.AllocHGlobal(OleObjectStorage.Length);
            Marshal.Copy(OleObjectStorage, 0, hGlobal, OleObjectStorage.Length);
            OLE32.CreateILockBytesOnHGlobal(hGlobal, false, out LockBytes);
            OLE32.IOleObject pOle = null;

            Ole.HRESULT result = OLE32.StgOpenStorageOnILockBytes(LockBytes, null, OLE32.STGM.STGM_READWRITE | OLE32.STGM.STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out RootStorage);

            System.Runtime.InteropServices.ComTypes.STATSTG MetaDataInfo;
            RootStorage.Stat(out MetaDataInfo, 0);
            ClsId = MetaDataInfo.clsid;
            OLE32.ProgIDFromCLSID(ref ClsId, out ProgID);

            Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

            Ole.IOleClientSite ole_cs = null;
            result = OLE32.OleLoad(RootStorage, ref IID_IUnknown, ole_cs, out pOle);
            if (result != Ole.HRESULT.S_OK)
            {
                int res = Marshal.GetLastWin32Error();
                return null;
            }

            IntPtr pUnknwn = Marshal.GetIUnknownForObject(pOle);
            result = OLE32.OleSetContainedObject(pUnknwn, true);
            result = OLE32.OleNoteObjectVisible(pUnknwn, true);
            Marshal.Release(pUnknwn);
            result = OLE32.OleRun(pOle);

            //tagSIZEL sizel = new tagSIZEL();
            //pOle.GetExtent(1, ref sizel);
            //float LogUnitsPerDevPixel_X = 1, LogUnitsPerDevPixel_Y = 1;

            //using (Graphics g = Graphics.FromImage(new Bitmap(1, 1)))
            //{
            //    HandleRef hdcsrc = new HandleRef(g, g.GetHdc());
            //    LogUnitsPerDevPixel_X = ((float)UNITS_PER_INCH) / ((float)GetDeviceCaps(hdcsrc.Handle, LOGPIXELSX));
            //    LogUnitsPerDevPixel_Y = ((float)UNITS_PER_INCH) / ((float)GetDeviceCaps(hdcsrc.Handle, LOGPIXELSY));
            //    g.ReleaseHdc(hdcsrc.Handle);
            //}

            Rectangle Size = AbsoluteImageArea;

            Rectangle rect = new Rectangle(0, 0, Size.Width * 3, Size.Height * 3);
            Bitmap m = new Bitmap(rect.Width, rect.Height);

            using (Graphics g = Graphics.FromImage(m))
            {
                Color bgColor = Color.White;


                bool MakeTransparent = false;
                if (ProgID == "BMP1C.Bmp1cCtrl.1")
                    if ((pOle as _DBmp_1c).GrMode == 1) // Иначе рисует только маску
                    {
                        (pOle as _DBmp_1c).GrMode = 2;
                        MakeTransparent = true;
                    }

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
                {
                    int ColorIndex = FormatCell.bBackground;
                    if (ColorIndex >= 0 || ColorIndex < a1CPallete.Length)
                        bgColor = Color.FromArgb((int)(a1CPallete[ColorIndex] + 0xFF000000));
                }
                else
                    MakeTransparent = true;

                g.Clear(bgColor);
                HandleRef hdcsrc = new HandleRef(g, g.GetHdc());
                result = OLE32.OleDraw(pOle, 1, hdcsrc, ref rect);
                g.ReleaseHdc(hdcsrc.Handle);
                if (MakeTransparent)
                    m.MakeTransparent(bgColor);
            }
            Marshal.ReleaseComObject(RootStorage);
            Marshal.ReleaseComObject(LockBytes);
            Marshal.FreeHGlobal(hGlobal);

            OLE32.CoUninitialize();
            return m;
        }

        private Image LoadPicture(BinaryReader br)
        {
            uint xz = br.ReadUInt32();
            int PictureSize = br.ReadInt32();
            byte[] pictureBuffer = br.ReadBytes(PictureSize);

            using (var memoryStream = new MemoryStream(pictureBuffer))
            {
                Bitmap Pic = Image.FromStream(memoryStream) as Bitmap;
                bool MakeTransparent = false;

                if (!FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
                    MakeTransparent = true;

                if (MakeTransparent)
                    Pic.MakeTransparent(Color.White);

                return Pic;
            }
        }
    }

}
