using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcellAPI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            SetContextMenu();
        }

        private void CopyAll()
        {
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.SelectAll();
            var dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app;
            Workbook workbook;
            Worksheet worksheet;
            Range range;

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
                    dataGridView1.Rows.Add(range.Cells[row, 1].Text, range.Cells[row, 2].Text,
                        range.Cells[row, 3].Text);
                }
                workbook.Close();
                app.Quit();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CopyAll();
            Microsoft.Office.Interop.Excel.Application app;
            Workbook workbook;
            Worksheet worksheet;
            object misValue = System.Reflection.Missing.Value;
            app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            workbook = app.Workbooks.Add(misValue);
            worksheet = (Worksheet)workbook.Worksheets.Item[1];
            var CR = (Range)worksheet.Cells[1, 1];
            CR.Select();
            worksheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);   
        }
        
        public void SetContextMenu()
        {
            var contextMenu = new ContextMenuStrip();
            var alignItem = new ToolStripMenuItem("Text Align");
            var colorItem = new ToolStripMenuItem("Color");
            var fontItem = new ToolStripMenuItem("Font");
            contextMenu.Items.AddRange(new ToolStripItem[] {alignItem, colorItem, fontItem});
            dataGridView1.ContextMenuStrip = contextMenu;
        }
    }
}