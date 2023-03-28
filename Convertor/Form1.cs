using ClosedXML.Excel;
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
using static System.Net.Mime.MediaTypeNames;

namespace Convertor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Title = "Введите имя файла";
            dialog.Filter = "Текстовые файлы|*.PRT|Все файлы|*.*";

            if (dialog.ShowDialog() != DialogResult.OK)
                return;

            var saveDialog = new SaveFileDialog();
            saveDialog.Title = "Введите имя файла для сохранения";
            saveDialog.Filter = "Файлы Excel|*.xlsx|Все файлы|*.*";

            

            if (saveDialog.ShowDialog() != DialogResult.OK)
                return;
            string resultFile = saveDialog.FileName;

            string text = dialog.FileName;


            string contents = File.ReadAllText(text);
            

            using var workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Sheet");
            worksheet.Cell("A1").Value = contents;
            if (!resultFile.EndsWith(".xslx", StringComparison.OrdinalIgnoreCase))
            resultFile += ".xlsx";
            workbook.SaveAs(resultFile);
                     




        }

     
    }
}

