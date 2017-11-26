using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.IO;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
            this.FormClosing += Form5_FormClosing;
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            progressBar1.Maximum = 10;
            progressBar1.Value = 0;
            progressBar1.Step = 2;
            if (PingTest() == true)
            {
                progressBar1.Value += progressBar1.Step;
                label1.Text += "可行";
            }
            else
            {
                MessageBox.Show("無法連線，請檢查網路狀態", "錯誤:ConnectError", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(0);
            }
            if (ExcelExist() == true)
            {
                progressBar1.Value += progressBar1.Step;
                label2.Text += "可行";
            }
            else
            {
                MessageBox.Show("Excel無法運作，請檢查Excel是否可以正常運作", "錯誤:ExcelRunError", MessageBoxButtons.OK,MessageBoxIcon.Error);
                Environment.Exit(0);
            }
            if (PasswordExist() == true)
            {
                progressBar1.Value += progressBar1.Step;
                label3.Text += "存在";
            }
            else
            {
                progressBar1.Value += progressBar1.Step;
                label3.Text += "不存在";
            }
            if (SeatListExist() == true)
            {
                progressBar1.Value += progressBar1.Step;
                label4.Text += "存在";
            }
            else
            {
                progressBar1.Value += progressBar1.Step;
                label4.Text += "不存在";
            }
            progressBar1.Value += progressBar1.Step;
            if (progressBar1.Value == 10)
            {
                Close();
            }
        }

        private void Form5_FormClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (progressBar1.Value != 10)
            {
                Environment.Exit(0);
            }
        }

        private Boolean PingTest()
        {
            Ping ping = new Ping();
            String google = "www.google.com.tw";
            PingReply reply;
            try
            {
                reply = ping.Send(google);
                if (reply.Status == IPStatus.Success)
                {
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private Boolean ExcelExist()
        {
            Excel.Application excel = new Excel.Application();
            if (excel != null)
            {
                return true;
            }
            return false;
        }

        private Boolean PasswordExist()
        {
            if(File.Exists(System.IO.Directory.GetCurrentDirectory() + @"\site.txt")){
                return true;
            }
            return false;
        }

        private Boolean SeatListExist()
        {
            if (File.Exists(System.IO.Directory.GetCurrentDirectory() + @"\SeatList.txt"))
            {
                return true;
            }
            return false;
        }
    }
}
