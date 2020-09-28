using System;
using System.IO;
using System.Threading.Tasks;
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
            OpenFileDialog dd = new OpenFileDialog {Filter = "Moxel |*.mxl", DefaultExt = "mxl"};
            if (dd.ShowDialog() != DialogResult.OK)
                return;
            var mdfilename = dd.FileName;
            Task.Run(() =>
                {
                    button1.Invoke(new Action(() => button1.Enabled = false));
                    Task.Run(() => ExcelWriter_onProgress(0));
                    ExcelWriter.onProgress += ExcelWriter_onProgress;
                    try
                    {
                        var mxl = new Moxel.Moxel(mdfilename);
                        mxl.SaveAs(Path.ChangeExtension(mdfilename, "xlsx"), Moxel.SaveFormat.Excel);
                    }
                    catch
                    {
                        label1.Invoke(new Action(() => label1.Text = "Ошибка"));
                    }
                    finally
                    {

                        ExcelWriter.onProgress -= ExcelWriter_onProgress;
                        Task.Run(() => ExcelWriter_onProgress(100));
                        button1.Invoke(new Action(() => button1.Enabled = true));
                    }
                }
            );
        }

        private void ExcelWriter_onProgress(int progress)
        {
            label1.Invoke(new Action(() => label1.Text = $"{progress} %"));
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
