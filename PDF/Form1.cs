using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text.pdf;
using VisioForge.MediaFramework.Helpers;

namespace PDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        public static string GetExportPDFPath()
        {
            string filename = System.DateTime.Now.ToString().Replace("/", "").Replace(":", "") + "XXX";
            var exportFileName = filename + ".pdf";
            var saveFileDialog = new SaveFileDialog { FileName = exportFileName };
            saveFileDialog.Filter = "匯出資料檔案(*.pdf)|*.pdf";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.Title = "匯出資料檔案";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                return saveFileDialog.FileName;
            }
            else
            {
                return String.Empty;
            }
        }
        public void ConvertJPGToPDF(string jpgfile, string pdf)
        {
            iTextSharp.text.Rectangle PageSize = new iTextSharp.text.Rectangle(this.Width, this.Height);
            var document = new iTextSharp.text.Document(PageSize, 25, 25, 25, 25);
            //var document = new iTextSharp.text.Document(new iTextSharp.text.Rectangle(), 25, 25, 25, 25);
            using (var stream = new FileStream(pdf, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                PdfWriter.GetInstance(document, stream);
                document.Open();
                using (var imageStream = new FileStream(jpgfile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var image = iTextSharp.text.Image.GetInstance(imageStream);
                    if (image.Height > PageSize.Height - 25)
                    {
                        image.ScaleToFit(PageSize.Width - 25, PageSize.Height - 25);
                    }
                    else if (image.Width > PageSize.Width - 25)
                    {
                        image.ScaleToFit(PageSize.Width - 25, PageSize.Height - 25);
                    }
                    image.Alignment = iTextSharp.text.Image.ALIGN_MIDDLE;
                    document.Add(image);
                }
                document.Close();
            }
        }

        private void btnExport_Click_1(object sender, EventArgs e)
        {
            //GDI+
            //this.groupBox1.Visible = false;
            //this.label1.Focus();
            Bitmap bmp = new Bitmap(this.Size.Width, this.Size.Height);
            this.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, Width, Height));
            bmp.Save(".//temp.jpg");
            //選擇匯出pdf的路徑並匯出 ITextSharp.dll外掛
            string pdfpath = GetExportPDFPath();
            if (pdfpath == String.Empty)
            {
                MessageBox.Show("匯出失敗，請重試!!");
                //this.groupBox1.Visible = true;
                return;
            }
            else
            {
                ConvertJPGToPDF(".//temp.jpg", pdfpath);
                MessageBox.Show("匯出成功!!");
                this.Close();
            }
        }
    }
}
