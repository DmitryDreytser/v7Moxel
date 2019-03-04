using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using AddIn;
using System.ComponentModel;
using System.Windows.Forms;

namespace Moxel
{
   /* 
    [ComVisible(true)]
    public enum Format
    {
        Excel = 1,
        Html =2,
        PDF = 3
    }

    //[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    //[Guid("9CDFF492-5B36-49EE-8AB4-4C21B2FD26C9")]
    //[ComVisible(true)]
    internal interface IConverter
    {
        void Attach(dynamic Table);
        void Save(Format format);
    }
    */
    /*
    [ComVisible(true)]
    [ComSourceInterfaces(typeof(ILanguageExtender), typeof(IInitDone))]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComDefaultInterface(typeof(ILanguageExtender))]
    [Description("Конвертер MOXEL")]
    [ProgId("AddIn.Moxel.Converter")]
    [Guid("B548768B-9589-4D90-9719-CDEF61A1BC5A")]
    public class Converter : AddIn.AddIn//, IConverter
    {

        public void Attach(dynamic Table)
        {
            MessageBox.Show("Attach");
            string tempfile = Path.GetTempFileName();
            Table.Записать(tempfile, 1);
        }

        public void Save(Format format)
        {
            MessageBox.Show("Save!");
        }

   }
*/    
}
