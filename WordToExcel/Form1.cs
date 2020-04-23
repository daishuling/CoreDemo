using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Collections;

namespace WordToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Document doc = new Document();
            doc.LoadFromFile(@"D:\图片\test.docx");

            //获取第一个文本框
            //TextBox textbox = document.TextBoxes[0];

            ////获取文本框中第一个表格
            //Table table = textbox.Body.Tables[0] as Table;

            //StringBuilder sb = new StringBuilder();

            ////遍历表格中的段落并提取文本
            //foreach (TableRow row in table.Rows)
            //{
            //    foreach (TableCell cell in row.Cells)
            //    {
            //        foreach (Paragraph paragraph in cell.Paragraphs)
            //        {
            //            sb.AppendLine(paragraph.Text);
            //        }
            //    }
            //}
            //File.WriteAllText("text.txt", sb.ToString());
            //初始化变量i
            int i = 0;
            List < Table > t=new List<Table>();
            //遍历文档中section
            foreach (Section section in doc.Sections)
            {
                //获取每一个section的表格数
                i = i + section.Tables.Count;
               // t.Add(doc.Sections[i].Tables[]);
            }
           //
            int c = i;
            //TableCollection table = section.Tables as TableCollection;
            //int count = table.Count;
            //Table t=table[0] as Table;
            //Table t1 = table[1] as Table;
            ////string s =;
            //string s2=t1.Rows[1].Cells[1].Paragraphs[0].Text;
        }
    }
}
