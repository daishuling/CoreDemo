using Saibaiwei;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IntoEngish change = new IntoEngish();
            Double s = Convert.ToDouble(textBox1.Text.ToString());
            textBox2.Text = change.NumberToString(s);
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //           /* 
            //*设置textBox只能输入数字（正数，负数，小数） 
            //*使用了TextBox的KeyPress事件
            //*/&& e.KeyChar != (char)('-')
            //           //允许输入数字、小数点、删除键和负号  
            if ((e.KeyChar < 48 || e.KeyChar > 57) && e.KeyChar != 8 && e.KeyChar != (char)('.') )
                {
                    MessageBox.Show("请输入正确的数字");
                    this.textBox1.Text = "";
                    e.Handled = true;
                }
                if (e.KeyChar == (char)('-'))
                {
                    if (textBox1.Text != "")
                    {
                        MessageBox.Show("请输入正确的数字");
                        this.textBox1.Text = "";
                        e.Handled = true;
                    }
                }
                /*小数点只能输入一次*/
                if (e.KeyChar == (char)('.') && ((TextBox)sender).Text.IndexOf('.') != -1)
                {
                    MessageBox.Show("请输入正确的数字");
                    this.textBox1.Text = "";
                    e.Handled = true;
                }
                /*第一位不能为小数点*/
                if (e.KeyChar == (char)('.') && ((TextBox)sender).Text == "")
                {
                    MessageBox.Show("请输入正确的数字");
                    this.textBox1.Text = "";
                    e.Handled = true;
                }
                /*第一位是0，第二位必须为小数点*/
                if (e.KeyChar != (char)('.') && ((TextBox)sender).Text == "0")
                {
                    MessageBox.Show("请输入正确的数字");
                    this.textBox1.Text = "";
                    e.Handled = true;
                }
                /*第一位是负号，第二位不能为小数点*/
                if (((TextBox)sender).Text == "-" && e.KeyChar == (char)('.'))
                {
                    MessageBox.Show("请输入正确的数字");
                    this.textBox1.Text = "";
                    e.Handled = true;
                }
            

        }
    }
}
