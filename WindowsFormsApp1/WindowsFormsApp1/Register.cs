using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using System.IO;

namespace WindowsFormsApp1
{
    public partial class Register : Form
    {
        public Register()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool a = checkUser(textBox1.Text);
            bool b = checkPassword(textBox2.Text);
            bool c = checkEmail(textBox4.Text);
            String users = KeySet(textBox1.Text);
            String passs = KeySet(textBox2.Text);
            if(a && b && c)
            {
                try
                {
                    using(FileStream fs = File.Create(System.IO.Directory.GetCurrentDirectory() + @"\site.txt"))
                    {

                        Byte[] user = new UTF8Encoding(true).GetBytes(users + ",");
                        Byte[] pass = new UTF8Encoding(true).GetBytes(passs + ",");
                        fs.Write(user, 0,user.Length);
                        fs.Write(pass, 0, pass.Length);
                        label8.Text = System.IO.Directory.GetCurrentDirectory() + @"\site.txt";
                    }
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                MessageBox.Show("已儲存並將Key寄到了信箱，請等待視窗關閉","Success",MessageBoxButtons.OK,MessageBoxIcon.Information);
                List<String> list = new List<string>();
                list.Add(textBox4.Text);
                SendMailBase(list, "-密鑰提醒-", "你好，歡迎使用，製作者:Coder_Sigtuna，您的密碼是：" + textBox2.Text);
                Close();
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private Boolean checkUser(String s)
        {
            if (s.Length > 6)
            {
                return true;
            }
            else
            {
                label3.Text = "帳號必須大於6個字";
                label3.ForeColor = Color.Red;
            }
            return false;
        }

        private Boolean checkPassword(String s)
        {
            if(s.Length > 6)
            {
                return true;
            }
            else
            {
                label9.Text = "密碼必須大於6個字";
                label9.ForeColor = Color.Red;
            }
            return false;
        }

        private Boolean checkEmail(String s)
        {
            if (s.Contains("@") && !s.Contains("@sgm."))
            {
                return true;
            }
            else if (s.Contains("@sgm."))
            {
                label10.Text = "請勿使用學校信箱(@後面為sgm)";
                label10.ForeColor = Color.Red;
            }
            else
            {
                label10.Text = "信箱格式不正確";
                label10.ForeColor = Color.Red;
            }
            return false;
        }

        private Boolean checkKey(String s)
        {
            if(s.Length == textBox2.Text.Length)
            {
                return true;
            }

            return false;
        }

        private String KeySet(String s)
        {
            byte[] bytes = System.Text.Encoding.GetEncoding("utf-8").GetBytes(s);
            return Convert.ToBase64String(bytes);
        }

        public static void SendMailBase(List<string> ReceivingMails, string MailSubject, string MailBody)
        {
            //謝謝勇者大大
            // 建立類別
            MailMessage Mail = new MailMessage();
            // 設定發件人
            Mail.From = new MailAddress("sigtuna910625@gmail.com", "Sigtuna's Secretary");
            // 加入收件人
            foreach (string ReceivingMail in ReceivingMails)
            { Mail.Bcc.Add(ReceivingMail); }
            // 等級
            Mail.Priority = MailPriority.High;
            // 標題
            Mail.Subject = MailSubject;
            // 內文
            Mail.Body = MailBody;
            Mail.IsBodyHtml = true;
            Mail.BodyEncoding = System.Text.Encoding.UTF8;
            // 發信
            SmtpClient SmtpServer = new SmtpClient();
            SmtpServer.Credentials = new System.Net.NetworkCredential(
                "sigtuna910625@gmail.com",
                "han910625");
            SmtpServer.Port = 587;
            SmtpServer.Host = "smtp.gmail.com";
            SmtpServer.EnableSsl = true;
            SmtpServer.Send(Mail);
            // 釋放
            Mail.Dispose();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
    }

}
