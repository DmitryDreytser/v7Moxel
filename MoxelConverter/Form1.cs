using System;
using System.IO;
using System.Windows.Forms;
using Moxel;

namespace MoxelConverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dd = new OpenFileDialog { Filter = "Moxel |*.mxl", DefaultExt = "mxl" };

            dd.DefaultExt = "*.mxl";
            if (dd.ShowDialog() != DialogResult.OK)
                return;
            string mdfilename = dd.FileName;
            Moxel.Moxel mxl = new Moxel.Moxel(mdfilename);

            mxl.SaveAs(Path.ChangeExtension(mdfilename, "xlsx"), Moxel.SaveFormat.Excel);
            
            mxl.SaveAs(Path.ChangeExtension(mdfilename, "pdf"), Moxel.SaveFormat.PDF);
            mxl.SaveAs(Path.ChangeExtension(mdfilename, "html"), Moxel.SaveFormat.Html);

            //Moxel.ExcelWriter.Save(mxl, mdfilename + ".xlsx");
            //Moxel.PDFWriter.Save(mxl, Path.ChangeExtension(mdfilename, "pdf"));
            //Moxel.HtmlWriter.Save(mxl, mdfilename + ".html");
            
        }
    }
}
