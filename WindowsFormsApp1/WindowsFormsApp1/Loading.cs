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
using System.Net;
using System.Diagnostics;
namespace WindowsFormsApp1
{
    public partial class Loading : Form
    {
        public Loading()
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
            Updater();
            OldVersionDelete();
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

        private void label6_Click(object sender, EventArgs e)
        {
            
        }
        private void Updater()
        {
            try
            {
                WebClient web = new WebClient();
                String version = web.DownloadString("https://pastebin.com/raw/xbf0H5N5");
                String now = "v0.035";
                if (version != now)
                {
                    DialogResult dr = MessageBox.Show("有最新版本: " + version + "\n請問是否更新?", "UpdateNotice", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (dr == DialogResult.Yes)
                    {
                        String path = System.IO.Directory.GetCurrentDirectory() + @"\" + "WindowsFormsApp1.exe";
                        String DownloadLink = "https://github.com/SigtunaDev/Check-IN-System/releases/download/" + version + @"/Check-In-System-by-Sigtuna.exe";
                        String FileName = "Check-IN_System_" + version + "_.exe";
                        web.DownloadFile(DownloadLink, FileName);
                        MessageBox.Show("下載成功!");
                        Process.Start("Check-IN_System_" + version + "_.exe");
                        Environment.Exit(0);
                    }
                }
            }
            catch (System.Net.WebException)
            {
                MessageBox.Show("檔案連結不存在，請聯繫作者", "-BUG_Warning-", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void OldVersionDelete()
        {
            String FileName = "Check-In-System_v0.02_.exe";
            String path = System.IO.Directory.GetCurrentDirectory() + @"\" + FileName;
            if (File.Exists(path))
            {
                MessageBox.Show("刪除舊版本檔案: Check-In-System-by-Sigtuna.exe\n提醒: 請勿更改檔案名稱", "DeleteOldVersion", MessageBoxButtons.OK, MessageBoxIcon.Information);
                File.Delete(path);
            }
        }
    }
}
