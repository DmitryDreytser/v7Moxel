using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Moxel;
using v7Moxel.Moxel.ExcelWriter;

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
                    ExcelWriter.OnProgress += ExcelWriter_onProgress;
                    var load = 0L;
                    var save = 0L;
                    try
                    {
                        var sw = Stopwatch.StartNew();
                        var mxl = new Moxel.Moxel(mdfilename);
                        load = sw.ElapsedMilliseconds;
                        sw.Restart();
                        mxl.SaveAs(Path.ChangeExtension(mdfilename, "xlsx"), Moxel.SaveFormat.Excel);
                        save = sw.ElapsedMilliseconds;
                        Task.Run(() => ExcelWriter_onProgress(100));
                    }
                    catch(Exception ex)
                    {
                        label1.Invoke(new Action(() => label1.Text = $"Ошибка {ex}"));
                    }
                    finally
                    {
                        ExcelWriter.OnProgress -= ExcelWriter_onProgress;
                        button1.Invoke(new Action(() => button1.Enabled = true));
                        label1.Invoke(new Action(() => label1.Text = $"Загрузка {load}мс\r\nСохранение {save}мс"));
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
