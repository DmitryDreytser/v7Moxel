using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Moxel.Moxel;

namespace Moxel
{
    public class DataCell
    {
        public Cellv6 FormatCell;
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
            FormatCell = br.Read<Cellv6>();

            if (Parent != null)
                if (Parent.Version == 7)
                    TextOrientation = br.ReadInt16();

            if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.Text))
            {
                Text = br.ReadCString();
            }

            if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.Value))
            {
                Value = br.ReadCString();
            }

            if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.Data))
            {
                Data = br.ReadBytes(br.ReadCount());
            }
        }

        public static implicit operator Cellv6(DataCell dc)
        {
            return dc.FormatCell;
        }

        public static implicit operator DataCell(Cellv6 c)
        {
            return new DataCell { FormatCell = c, TextOrientation = 0, Data = new byte[0], Parent = null };
        }

    }
}
