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
        private Api _excel = new Api();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) => Api.ImportFile(dataGridView1);

        private void button2_Click(object sender, EventArgs e) => Api.ExportFile(dataGridView1);

        private void button3_Click(object sender, EventArgs e)
        {
            Api.CalcAverage(dataGridView1);
        }
    }
}