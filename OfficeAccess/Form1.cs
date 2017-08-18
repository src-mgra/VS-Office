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
using Word = Microsoft.Office.Interop.Word;

namespace OfficeAccess
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //Excel aufrufen
            var excelApp = new Excel.Application();

            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet worksheet = (Excel._Worksheet)excelApp.ActiveSheet;
            
            worksheet.Cells[1,1]="Wert1";

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Word
            var wordApp = new Word.Application();

            wordApp.Visible = true;
            wordApp.Documents.Add();        
  
            wordApp.Selection.InsertAfter("Test Word");
        }
    }
}
