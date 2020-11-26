using System;
using System.Net;
using System.Windows.Forms;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellAPI
{
    public class Api
    {
        private const string Url = "http://localhost:3000/posts";

        private static void CopyAll(DataGridView dataGrid)
        {
            dataGrid.RowHeadersVisible = false;
            dataGrid.SelectAll();
            var dataObj = dataGrid.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        public static void ImportFile(DataGridView dataGrid)
        {
            dataGrid.Rows.Clear();
            Microsoft.Office.Interop.Excel.Application app;
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            Excel.Range range;

            var dialog = new OpenFileDialog();
            var path = dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : "";

            if (path != "")
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                workbook = app.Workbooks.Open(path);
                worksheet = workbook.Worksheets["main"];
                range = worksheet.UsedRange;
                for (var row = 2; row <= range.Rows.Count; row++)
                {
                    dataGrid.Rows.Add(range.Cells[row, 1].Text, range.Cells[row, 2].Text,
                        range.Cells[row, 3].Text);
                }
            }
        }

        public static void ExportFile(DataGridView dataGrid, bool isCalculated = false)
        {
            CopyAll(dataGrid);
            Microsoft.Office.Interop.Excel.Application app;
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            object misValue = System.Reflection.Missing.Value;
            app = new Microsoft.Office.Interop.Excel.Application() /*{Visible = true}*/;
            workbook = app.Workbooks.Add(misValue);
            worksheet = (Excel.Worksheet) workbook.Worksheets.Item[1];
            var select = (Excel.Range) worksheet.Cells[1, 1];
            select.Select();
            if (isCalculated)
            {
                worksheet.Cells[dataGrid.Rows.Count + 1, 1].FormulaLocal = "=СРЗНАЧ(;A2:A11)";
            }

            worksheet.PasteSpecial(select, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, true);
            app.Visible = true;
            dataGrid.ClearSelection();
        }

        public static void CalcAverage(DataGridView dataGrid)
        {
            dataGrid.Rows.Clear();
            var data = GetData();
            foreach (var col in data)
            {
                dataGrid.Rows.Add(col.age, col.name, col.text);
            }

            ExportFile(dataGrid, true);
        }

        private static dynamic GetData()
        {
            try
            {
                var client = new WebClient();
                var data = client.DownloadString(Url);
                return JsonConvert.DeserializeObject<dynamic>(data);
            }
            catch
            {
                return "";
            }
        }
    }
}