using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net.Mail;

namespace WindowsFormsApp1
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            this.FormClosing += formClose;
        }

        public bool b = false;
        public bool unlock = false;
        public bool hidden = true;
        public void Form3_Load(object sender, EventArgs e)
        {
            Form3 f3 = new WindowsFormsApp1.Form3();
        }
        public Boolean unLock()
        {
            unlock = true;
            if (b == true)
            {
                return true;
            }
            return false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (unlock == false)
            {
                if (!File.Exists(System.IO.Directory.GetCurrentDirectory() + @"\site.txt"))
                {
                    if (textBox1.Text == "root" && textBox2.Text == "root")
                    {
                        MessageBox.Show("在等等的頁面，請儲存一個新密碼","提示:註冊",MessageBoxButtons.OK,MessageBoxIcon.Information);
                        Form4 f4 = new Form4();
                        f4.Show();
                    }
                    else
                    {
                        MessageBox.Show("您還沒有註冊任何帳號，請在帳號和密碼輸入root以進行註冊", "提示:註冊", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    string text = File.ReadAllText(System.IO.Directory.GetCurrentDirectory() + @"\site.txt");
                    String user = text.Split(',')[0];
                    String password = text.Split(',')[1];
                    String ruser = System.Text.Encoding.GetEncoding("utf-8").GetString(Convert.FromBase64String(user));
                    String rpassword = System.Text.Encoding.GetEncoding("utf-8").GetString(Convert.FromBase64String(password));
                    if (textBox1.Text == ruser && textBox2.Text == rpassword)
                    {
                        b = true;
                        MessageBox.Show("登入成功，歡迎回來","welcome back",MessageBoxButtons.OK,MessageBoxIcon.Information);
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("帳號或密碼錯誤","wrong account",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                label1.Text += " LOCKED";
                string text = File.ReadAllText(System.IO.Directory.GetCurrentDirectory() + @"\site.txt");
                String user = text.Split(',')[0];
                String password = text.Split(',')[1];
                String ruser = System.Text.Encoding.GetEncoding("utf-8").GetString(Convert.FromBase64String(user));
                String rpassword = System.Text.Encoding.GetEncoding("utf-8").GetString(Convert.FromBase64String(password));
                if (textBox1.Text == ruser && textBox2.Text == rpassword)
                {
                    b = true;
                    MessageBox.Show("登入成功，歡迎回來");
                    this.Close();
                }
                else
                {
                    MessageBox.Show("帳號或密碼錯誤");
                }
            }
        }

        private void formClose(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (b == false && unlock == false)
            {
                Environment.Exit(0);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form3_Load()
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (button2.Text == "顯示密碼")
            {
                textBox2.UseSystemPasswordChar = false;
                button2.Text = "隱藏密碼";
            }
            else
            {
                textBox2.UseSystemPasswordChar = true;
                button2.Text = "顯示密碼";
            }
        }
    }
}
