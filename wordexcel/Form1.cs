
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace wordexcel
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

        private void KillProcess(string processName)
        {
            System.Diagnostics.Process myProc = new System.Diagnostics.Process();
            try
            {
                foreach (System.Diagnostics.Process thisProc in System.Diagnostics.Process.GetProcessesByName(processName))
                {
                    thisProc.Kill();
                }
            }
            catch (Exception exc)
            {
                throw new Exception("", exc);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            KillProcess("excel");
            OpenFileDialog oFile = new OpenFileDialog();
            oFile.Filter = @"Word文件(*.doc,*.docx)|*.doc;*.docx";
            if (oFile.ShowDialog() == DialogResult.OK)
            {
                object oFileName = oFile.FileName;
                //object oFileName = @"D:\图片\test.docx";
                object oReadOnly = true;
                object oMissing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word._Document oDoc;
                Microsoft.Office.Interop.Word.Application oWord= new Microsoft.Office.Interop.Word.Application();
                oWord.Visible = true;//只是为了方便观察
                oDoc = oWord.Documents.Open(ref oFileName, ref oMissing, ref oReadOnly, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                string tableMessage="";
                string tmp;
                string path = AppDomain.CurrentDomain.BaseDirectory + "生成的excel文件";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string oExcelName = path + "\\" + "生成的excel" + ".xlsx";
                object oRead = false;
                object oMiss = System.Reflection.Missing.Value;
                var oExcel = new Microsoft.Office.Interop.Excel.Application();
                //Microsoft.Office.Interop.Excel.Workbook xBook = oExcel.Workbooks.Open(oExcelName, oMiss, oRead, oMiss, oMiss, oMiss, oMiss, oMiss, oMiss, oMiss, oMiss, oMiss,oMiss, oMiss, oMiss);
                Microsoft.Office.Interop.Excel.Workbook xBook = oExcel.Workbooks.Add(Type.Missing);

                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)xBook.ActiveSheet;
                //Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)xBook.Sheets.get_Item(oDoc.Tables.Count);
                // Microsoft.Office.Interop.Excel._Worksheet xSt1 = (Microsoft.Office.Interop.Excel._Worksheet)xBook.Sheets.Add(Missing.Value, ws, Missing.Value, Missing.Value);
                if (oDoc.Tables.Count <= 0)
                {
                    return;
                }
                if (oDoc.Tables.Count > 1)
                {
                    ws = xBook.Sheets.Add(Missing.Value, ws, oDoc.Tables.Count - 1, Missing.Value);
                }

                Microsoft.Office.Interop.Excel.Sheets shs = oExcel.Sheets;
                for (int tablePos = 1; tablePos <= oDoc.Tables.Count; tablePos++)
                {

                    if (oDoc.Tables.Count <= 1)
                    {
                    }
                    else
                    {
                        ws = xBook.Sheets[tablePos]; ;
                    }
                    Microsoft.Office.Interop.Word.Table nowTable = oDoc.Tables[tablePos];
                    int n = oDoc.Tables[tablePos].Rows.Count;
                    for (int rowPos = 1; rowPos <= nowTable.Rows.Count; rowPos++)
                    {

                        for (int columPos = 1; columPos <= nowTable.Columns.Count; columPos++)
                        {
                            Microsoft.Office.Interop.Excel.Range rang = (Microsoft.Office.Interop.Excel.Range)ws.Cells[rowPos, columPos];
                            //if ((bool)rang.MergeCells==false)
                            //{
                            tableMessage = nowTable.Cell(rowPos, columPos).Range.Text.Trim();
                            tmp = tableMessage.Replace("\r\a", "");
                            ws.Cells[rowPos, columPos] = tmp;
                            //}

                        }

                    }

                }
                ws.SaveAs(oExcelName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                MessageBox.Show(oExcelName + @"已成功生成" + oDoc.Tables.Count.ToString() + @"个表格");
                oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
                oExcel.Quit();

            }
    }
    }
}
