using System;
using System.Net;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace ExcellAPI
{
    public class API
    {
        private string url = "http://localhost:3000/posts";
        private void CopyAll(DataGridView dataGrid)
        {
            dataGrid.RowHeadersVisible = false;
            dataGrid.SelectAll();
            var dataObj = dataGrid.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        public void ImportFile(DataGridView dataGrid)
        {
            Microsoft.Office.Interop.Excel.Application app;
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;
            Microsoft.Office.Interop.Excel.Range range;

            var dialog = new OpenFileDialog();
            var path = dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : "";

            if (path != "")
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                workbook = app.Workbooks.Open(path);
                worksheet = workbook.Worksheets["main"];
                range = worksheet.UsedRange;
                var i = 0;
                for (var row = 2; row <= range.Rows.Count; row++)
                {
                    i++;
                    dataGrid.Rows.Add(range.Cells[row, 1].Text, range.Cells[row, 2].Text,
                        range.Cells[row, 3].Text);
                }
                workbook.Close();
                app.Quit();
            }
        }

        public void ExportFile(DataGridView dataGrid)
        {
            CopyAll(dataGrid);
            Microsoft.Office.Interop.Excel.Application app;
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;
            object misValue = System.Reflection.Missing.Value;
            app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            workbook = app.Workbooks.Add(misValue);
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Item[1];
            var select = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, 1];
            select.Select();
            worksheet.PasteSpecial(select, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }

        public void GetData(RichTextBox text)
        {
            var client = new WebClient();
            var data = client.DownloadString(url);
            dynamic dobj = JsonConvert.DeserializeObject<dynamic>(data);
            text.Text = dobj.ToString();
        }
    }
}