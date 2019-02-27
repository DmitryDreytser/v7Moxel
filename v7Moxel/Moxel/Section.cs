using System.IO;

namespace Moxel
{
    public class Section
    {
        int Begin;
        int End;
        int Level;
        string Name;

        public Section()
        { }

        public Section(BinaryReader br)
        {
            Begin = br.ReadInt32();
            End = br.ReadInt32();
            Level = br.ReadInt32();
            Name = br.ReadCString();
        }
    }
}
