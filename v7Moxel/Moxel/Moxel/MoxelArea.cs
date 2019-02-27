using System.IO;
using System.Runtime.InteropServices;

namespace Moxel
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
    public struct Area
    {
        public int Unknown1; // always 1
        public int Unknown2; // garbage
        public int AreaType;
        public int ColumnBegin;
        public int RowBegin;
        public int ColumnEnd;
        public int RowEnd;
    };

    public class MoxelArea
    {
        public string Name;
        public Area Area;

        public MoxelArea()
        { }

        public MoxelArea(BinaryReader br)
        {
            Name = br.ReadCString();
            Area = br.Read<Area>();
        }
    }

}
