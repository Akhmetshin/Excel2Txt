using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Txt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.RightToLeft = RightToLeft.Yes;
                textBox1.Text = ofd.FileName;
                string nameTxt = textBox1.Text + ".txt";

                Excel.Application ObjExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook;
                Excel.Worksheet ObjWorkSheet;

                ObjWorkBook = ObjExcel.Workbooks.Add(textBox1.Text);
                ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                ObjWorkBook.SaveAs(nameTxt,FileFormat.)
            }
        }
    }
}
