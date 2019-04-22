using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static Moxel.Moxel;

namespace Moxel
{
    [Serializable]
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public class DataCell
    {
        public CSheetFormat FormatCell { get; set; }
        
        //{
        //    get
        //    {
        //        return Parent.FormatsList[formatKey];
        //    }

        //    set
        //    {
        //            formatKey = value.GetHashCode();// + value.BgColor.GetHashCode() + value.wWidth;

        //            if (!Parent.FormatsList.ContainsKey(formatKey))
        //                Parent.FormatsList[formatKey] = value;
        //    }
        //}
    
        int formatKey;
        public string Text;
        public string Value;
        public byte[] Data;
        public short TextOrientation = 0;
        public Moxel Parent;

        public DataCell()
        { }
        public DataCell(BinaryReader br, Moxel parent = null)
        {
            Parent = parent;

            var tempFormat = br.Read<CSheetFormat>();

            if (Parent != null)
                if (Parent.Version == 7)
                    TextOrientation = br.ReadInt16();

            if (tempFormat.dwFlags.HasFlag(MoxelCellFlags.Text))
            {
                Text = br.ReadCString();
            }

            if (tempFormat.dwFlags.HasFlag(MoxelCellFlags.Value))
            {
                Value = br.ReadCString();
            }

            if (tempFormat.dwFlags.HasFlag(MoxelCellFlags.Data))
            {
                Data = br.ReadBytes(br.ReadCount());
            }

            FormatCell = tempFormat;
        }

        public DataCell(CSheetFormat formatCell, Moxel parent)
        {
            Parent = parent;
            FormatCell = formatCell;
        }

        public static implicit operator CSheetFormat(DataCell dc)
        {
            return dc.FormatCell;
        }

        //public static implicit operator DataCell(CSheetFormat c)
        //{
        //    DataCell dc = new DataCell();
        //    dc.FormatCell = c;
        //    dc.TextOrientation = 0;
        //    dc.Data = new byte[0];
        //    dc.Text = null;
        //    dc.Value = null;
        //    dc.Parent = null;
        //    return dc;
        //}

    }
}
