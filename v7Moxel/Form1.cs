using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace v7Moxel
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
            if (dd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;
            string mdfilename = dd.FileName;
            byte[] buffer = File.ReadAllBytes(mdfilename);

            Moxel.Moxel mxl = new Moxel.Moxel();
            mxl.Load(buffer);
            Moxel.HtmlConverter.SaveToHtml(mxl, mdfilename + ".html");
        }
    }
}
