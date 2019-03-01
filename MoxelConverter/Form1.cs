using System;
using System.Windows.Forms;

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

            mxl.SaveAs(mdfilename + ".html", Moxel.SaveFormat.Excel);

            //Moxel.ExcelWriter.Save(mxl, mdfilename + ".xlsx");
            //Moxel.PDFWriter.Save(mxl, mdfilename + ".pdf");
            //Moxel.HtmlWriter.Save(mxl, mdfilename + ".html");
            
        }
    }
}
