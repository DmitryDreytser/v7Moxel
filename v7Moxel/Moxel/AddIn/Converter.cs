using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using AddIn;

namespace Moxel
{
    public enum Format
    {
        Excel = 1,
        Html =2,
        PDF = 3
    }
    public interface IConverter
    {
        void Attach(dynamic Table);
        void Save(Format format);
    }

    [ProgId("Moxel.Converter")]
    [ComVisible(true)]
    class Converter : AddIn.AddIn, IConverter
    {
        public void Attach(dynamic Table)
        {
            string tempfile = Path.GetTempFileName();
            Table.Записать(tempfile, 1);
        }

        public void Save(Format format)
        {
            throw new NotImplementedException();
        }
    }
}
