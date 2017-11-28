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
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Net;
using System.Threading;
using Microsoft.Win32;
namespace WindowsFormsApp1
{
    public partial class MainPage : Form
    {

        Label[,] arraysH;
        String[] arraysS = new String[49];
        String wpath = System.IO.Directory.GetCurrentDirectory() + @"\SeatList.txt",excelpath;
        Label Glabel,Clabel;

        public MainPage(){
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
            Form1_Load();
        }

        private void Form1_Load()
        {
            timer1.Interval = 100;
            timer1.Enabled = true;
            Loader();
            Login();
            this.FormClosing += paintcolor;
            SystemEvents.TimeChanged += timeChange;
            this.timer1.Tick += new System.EventHandler(this.timeChange);
            ArrayInit();
            ClickEvent();
            SeatInit();
            SetTime();
            colorInit();
            VersionsUpdate();
        }

        private void timeChange(object sender, EventArgs e)
        {
            label65.Text = "現在時間: " + DateTime.Now.Hour + "點" + DateTime.Now.Minute + "分" + DateTime.Now.Second + "秒";
        } 

        public void VersionsUpdate()
        {
            WebClient web = new WebClient();
            String result = web.DownloadString("https://pastebin.com/raw/xbf0H5N5");
            label64.Text += result;
        }
        public void Loader()
        {
            Loading f5 = new Loading();
            f5.ShowDialog();
            f5.TopMost = true;
        }

        public Boolean ExcelExist()
        {
            Excel.Application excel = new Excel.Application();
            if (excel != null)
            {
                return true;
            }
            return false;
        }

        public void Login()
        {
            LoginPage f3 = new LoginPage();
            f3.ShowDialog();
            f3.TopMost = true;
        }

        public void SetTime()
        {
            label54.Text = DateTime.Now.ToString("yyyy MMM dd號 ddd");
        }

        public void ClickEvent()
        {
            for (int i = 0; i < 7; i++)
            {
                for (int j = 0; j < 7; j++)
                {
                    arraysH[i, j].Click += new EventHandler(labelClick);
                    arraysH[i, j].Tag = String.Format("{0},{1}", i + 1, j + 1);
                }
            }
        }
        Dictionary<int, String> map = new Dictionary<int, String>(); //儲存座標
        public void SeatInit()
        {
            Init();
            if (File.Exists(wpath))
            {
                string[] text = System.IO.File.ReadAllLines(wpath);
                map.Clear();
                for (int i = 0; i < text.Length; i++)
                {
                    String[] split = text[i].Split(',');
                    int h = Convert.ToInt16(split[0]);
                    int d = Convert.ToInt16(split[1]);
                    String name = split[2].ToString();
                    arraysH[h, d].Text = name.Split(' ')[0] + "\n" + name.Split(' ')[1];
                    arraysH[h, d].Tag = String.Format("{0},{1}", h + 1, d + 1);
                    map.Add(Convert.ToInt16(name.Split(' ')[0]), h + "," + d);
                }
                label58.Text = "SeatInit";
                for (int i = 0; i < 7; i++)
                {
                    for (int j = 0; j < 7; j++)
                    {
                        if (arraysH[i, j].Text == null || arraysH[i, j].Text == " ")
                        {
                            arraysH[i, j].Tag = "null";
                            arraysH[i, j].BackColor = Color.Gray;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("您還沒有設置任何的座位表，請按下方座位表更新去製作座位", "座位表初始化", MessageBoxButtons.OK, MessageBoxIcon.Information);
                for (int i = 0; i < 7; i++)
                {
                    for (int j = 0; j < 7; j++)
                    {
                        if (arraysH[i, j].Text == null || arraysH[i, j].Text == " ")
                        {
                            arraysH[i, j].Tag = "null";
                            arraysH[i, j].BackColor = Color.Gray;
                        }
                    }
                }
            }
        }

        public void paintcolor(object sender, System.ComponentModel.CancelEventArgs e)
        {
            String path = System.IO.Directory.GetCurrentDirectory() + @"\color.txt";
            if(File.Exists(path)){
                File.Delete(path);
            }
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(path))
            {
                Dictionary<String, Color> cmap = new Dictionary<String, Color>();
                for (int i = 0; i < map.Count(); i++)
                {
                    String h = map.ElementAt(i).Value.Split(',')[0];
                    String d = map.ElementAt(i).Value.Split(',')[1];
                    int r = Convert.ToInt16(h);
                    int t = Convert.ToInt16(d);
                    Color color = arraysH[r, t].BackColor;
                    cmap.Add(h + "," + d, color);
                }
                file.WriteLine(DateTime.Now.Year + "/" + DateTime.Now.Month + "/" + DateTime.Now.Day);
                for (int i = 0; i < cmap.Count(); i++)
                {
                    if (cmap.ElementAt(i).Value.ToKnownColor() != KnownColor.Yellow)
                    {
                        file.WriteLine(cmap.ElementAt(i).Key.ToString() + "_" + cmap.ElementAt(i).Value.ToKnownColor());
                    }
                    else
                    {
                        file.WriteLine(cmap.ElementAt(i).Key.ToString() + "_" + Color.Empty.ToKnownColor());
                    }
                }
            }
        }

        public void colorInit()
        {
            String path = System.IO.Directory.GetCurrentDirectory() + @"\color.txt";
            if (File.Exists(path))
            {
                String[] str = File.ReadAllLines(path);
                DateTime dt = DateTime.Now;
                if (str.ElementAt(0).Equals(dt.Year + "/" + dt.Month + "/" + dt.Day))
                {
                    for (int i = 1; i < str.Length; i++)
                    {
                        String[] sstr = str.ElementAt(i).Split('_');
                        String[] ssstr = sstr[0].Split(',');
                        int h = Convert.ToInt16(ssstr[0]);
                        int d = Convert.ToInt16(ssstr[1]);
                        arraysH[h, d].BackColor = Color.FromName(sstr[1]);
                    }
                }
                else
                {
                    File.Delete(path);
                }
            }
        }

        public void Init()
        {
            for(int i = 0; i < 7; i++)
            {
                for(int j = 0; j < 7; j++)
                {
                    arraysH[i, j].Text = " ";
                    arraysH[i, j].BackColor = DefaultBackColor;
                    arraysH[i, j].Tag = "";
                }
            }
        }

        private void ArrayInit()
        {
            arraysH = new Label[7, 7]
            {{label43,label36,label29,label22,label15,label7 ,label8 },
             {label44,label37,label30,label23,label16,label9 ,label5 },
             {label45,label38,label31,label24,label17,label10,label6 },
             {label46,label39,label32,label25,label18,label11,label3 },
             {label47,label40,label33,label26,label19,label12,label4 },
             {label48,label41,label34,label27,label20,label13,label2 },
             {label49,label42,label35,label28,label21,label14,label1 }};
        }

        private void labelClick(object sender, EventArgs e)
        {
            Label label = (Label)sender;
            if(label != Clabel && Clabel != null && Clabel.BackColor == Color.Yellow)
            {
                Clabel.BackColor = DefaultBackColor;
            }
            Glabel = label;
            if (Convert.ToString(label.Tag) != "null")
            {
                label56.Text = String.Format("第{0}排第{1}個", Convert.ToInt16(label.Tag.ToString().Split(',')[0]), Convert.ToInt16(label.Tag.ToString().Split(',')[1]));
            }
            else
            {
                label56.Text = null;
            }
            if (label.BackColor != Color.Empty && label.BackColor == DefaultBackColor)
            {
                label.BackColor = Color.Yellow;
                Clabel = label;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (Glabel != null && Glabel.Tag.ToString() != "null")
            {
                Glabel.BackColor = Color.LightGreen;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            NameTagUpload set = new NameTagUpload();
            set.ShowDialog();
            SeatInit();
        }

        private void Start()
        {
            for (int i = 0; i < arraysS.Length; i++)
            {
                String[] s = arraysS[i].Split(',');
                int r = Convert.ToInt16(s[1]);
                int p = Convert.ToInt16(s[2]);
                if (arraysH[r, p].Text != "")
                {
                    arraysH[r, p].Text = s[0];
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            SeatInit();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if(Glabel != null && Glabel.Tag.ToString() != "null")
            Glabel.BackColor = Color.Salmon;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Glabel != null && Glabel.Tag.ToString() != "null")
            Glabel.BackColor = Color.Goldenrod;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (Glabel != null && Glabel.Tag.ToString() != "null")
            Glabel.BackColor = Color.Pink;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (Glabel != null && Glabel.Tag.ToString() != "null")
                Glabel.BackColor = Color.Plum;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (Glabel != null && Glabel.Tag.ToString() != "null")
                Glabel.BackColor = Color.LightSkyBlue;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            Dictionary<int, Color> nmap = new Dictionary<int, Color>();
            Dictionary<int, String> em = new Dictionary<int, String>();
            for (int i = 0; i < map.Count; i++)
            {
                nmap.Add(map.ElementAt(i).Key,arraysH[Convert.ToInt16(map.ElementAt(i).Value.Split(',')[0]),Convert.ToInt16(map.ElementAt(i).Value.Split(',')[1])].BackColor);
            }
            for (int i = 0; i < nmap.Count; i++)
            {
                if (nmap.ElementAt(i).Value == Color.LightGreen)
                {
                    em.Add(nmap.ElementAt(i).Key, "出席");
                }
                if (nmap.ElementAt(i).Value == Color.Goldenrod)
                {
                    em.Add(nmap.ElementAt(i).Key, "病假");
                }
                if (nmap.ElementAt(i).Value == Color.Pink)
                {
                    em.Add(nmap.ElementAt(i).Key, "事假");
                }
                if (nmap.ElementAt(i).Value == Color.Plum)
                {
                    em.Add(nmap.ElementAt(i).Key, "公假");
                }
                if (nmap.ElementAt(i).Value == Color.LightSkyBlue)
                {
                    em.Add(nmap.ElementAt(i).Key, "遲到");
                }
                if (nmap.ElementAt(i).Value == Color.Salmon)
                {
                    em.Add(nmap.ElementAt(i).Key, "早退");
                }
                if (nmap.ElementAt(i).Value == DefaultBackColor)
                {
                    em.Add(nmap.ElementAt(i).Key, "未知");
                }
                if(nmap.ElementAt(i).Value == Color.MediumAquamarine)
                {
                    em.Add(nmap.ElementAt(i).Key, "曠課");
                }
            }
            int x = dayofweek(DateTime.Now.ToString("ddd"));
            for (int i = 0; i < em.Count; i++)
            {
                Console.WriteLine(em.ElementAt(i).Key + "," + em.ElementAt(i).Value);
            }
            ExcelCreate(label66.Text, x, em);
            label60.Text = DateTime.Now.TimeOfDay.Hours + "時" + DateTime.Now.TimeOfDay.Minutes + "分" + DateTime.Now.TimeOfDay.Seconds + "秒" + DateTime.Now.TimeOfDay.Milliseconds;
        }

        public int dayofweek(String s)
        {
            switch (s)
            {
                case "週一":
                    return 2;
                case "週二":
                    return 3;
                case "週三":
                    return 4;
                case "週四":
                    return 5;
                case "週五":
                    return 6;
                case "週六":
                    return 7;
                case "週日":
                    return 8;
                default:
                    return 9;
            }
        }

        Excel.Application excelapp;
        Excel._Workbook wBook;
        Excel._Worksheet wSheet;

        public void ExcelCreate(String path,int day,Dictionary<int, String> em)
        {
            if (File.Exists(path))
            {
                excelapp = new Excel.Application();
                excelapp.Visible = true;
                excelapp.DisplayAlerts = false;
                wBook = excelapp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
                DateTime date = DateTime.Now;
                Calendar cal = dfi.Calendar;
                bool found = false;
                foreach (Excel.Worksheet sheet in wBook.Sheets)
                {
                    if (sheet.Name == ("week" + cal.GetWeekOfYear(date, dfi.CalendarWeekRule, dfi.FirstDayOfWeek)))
                    {
                        found = true;
                    }
                }

                if (found == true)
                {
                    wSheet = wBook.Sheets[("week" + cal.GetWeekOfYear(date, dfi.CalendarWeekRule, dfi.FirstDayOfWeek))];
                    wSheet.Activate();
                }
                else
                {
                    Excel.Worksheet newWorksheet;
                    newWorksheet = (Excel.Worksheet)wBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    newWorksheet.Name = ("week" + cal.GetWeekOfYear(date, dfi.CalendarWeekRule, dfi.FirstDayOfWeek));
                    wSheet = newWorksheet;
                    wSheet.Activate();
                    NameTagInit(excelapp);
                }
                Excel.Range rng = excelapp.Range[excelapp.Cells[1, day], excelapp.Cells[em.Count + 2, day]];
                rng.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Empty);
                excelapp.get_Range("A1", "J99").Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excelapp.Cells[1, day] = DateTime.Now.ToString("ddd");
                //1,2 = 週一，1,3 = 週二 ...
                for (int i = 2; i < em.Count + 2; i++)
                {
                    if (em.ElementAt(i - 2).Value.Equals("出席"))
                    {
                        excelapp.Cells[em.ElementAt(i - 2).Key + 1, day] = em.ElementAt(i - 2).Value;
                    }else if(em.ElementAt(i-2).Value.Equals("遲到")){
                        Excel.Range rang = excelapp.Range[excelapp.Cells[em.ElementAt(i - 2).Key + 1, day], excelapp.Cells[em.ElementAt(i - 2).Key + 1, day]];
                        rang.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                        excelapp.Cells[em.ElementAt(i - 2).Key + 1, day] = "出席";
                    }
                    else
                    {
                        Excel.Range rang = excelapp.Range[excelapp.Cells[em.ElementAt(i - 2).Key + 1, day], excelapp.Cells[em.ElementAt(i - 2).Key + 1, day]];
                        rang.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        excelapp.Cells[em.ElementAt(i - 2).Key + 1, day] = em.ElementAt(i - 2).Value;
                    }
                }
                try
                {
                    wBook.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Console.WriteLine("儲存文件於 " + Environment.NewLine + path);
                }
                catch
                {
                    MessageBox.Show("儲存檔案出錯，檔案可能正在使用","錯誤:檔案無法儲存",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }



                wBook.Close(false, Type.Missing, Type.Missing);
                excelapp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp);
                wBook = null;
                wSheet = null;
                excelapp = null;
                GC.Collect();
                Console.Read();

            }
            else
            {
                MessageBox.Show("檔案不存在，請確認路徑", "錯誤:路徑", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "試算表|*.xlsx";
            ofd.ShowDialog();
            label66.Text = ofd.InitialDirectory + ofd.FileName;
            excelpath = ofd.InitialDirectory + ofd.FileName;
        }

        private void label60_Click(object sender, EventArgs e)
        {

        }

        private void NameTagInit(Excel.Application excelapp)
        {
            String path = System.IO.Directory.GetCurrentDirectory() + @"\NameTag.txt";
            String[] list = File.ReadAllLines(path);
            for (int i = 2; i < list.Count() + 2; i++)
            {
                excelapp.Cells[i, 1] = list[i-2]; 
            }
        }
        bool unlockb = false;
        private void button9_Click(object sender, EventArgs e)
        {
            if (button9.Text.Equals("鎖定"))
            {
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;
                button7.Enabled = false;
                button12.Enabled = false;
                button5.Enabled = false;
                button8.Enabled = false;
                button6.Enabled = false;
                button10.Enabled = false;
                button14.Enabled = false;
                button15.Enabled = false;
                button18.Enabled = false;
                button9.Text = "解鎖";
                label61.Text = "已鎖定 LOCKET";
                label62.Visible = true;
                label61.ForeColor = Color.Red;
                textBox2.Visible = true;
                button13.Visible = true;
                textBox2.Text = "";
            }
            else if (button9.Text.Equals("解鎖"))
            {
                unlock();
                if (unlockb == true)
                {
                    button1.Enabled = true;
                    button2.Enabled = true;
                    button3.Enabled = true;
                    button4.Enabled = true;
                    button7.Enabled = true;
                    button12.Enabled = true;
                    button5.Enabled = true;
                    button8.Enabled = true;
                    button6.Enabled = true;
                    button10.Enabled = true;
                    button14.Enabled = true;
                    button15.Enabled = true;
                    button18.Enabled = true;
                    label61.Text = "";
                    label61.ForeColor = Color.Empty;
                    button9.Text = "鎖定";
                    textBox2.Visible = false;
                    label62.Visible = false;
                    button13.Visible = false;
                }
            }
            else
            {
                MessageBox.Show("debug");
            }
        }

        public void unlock()
        {
            string text = File.ReadAllText(System.IO.Directory.GetCurrentDirectory() + @"\site.txt");
            String user = text.Split(',')[0];
            String password = text.Split(',')[1];
            String ruser = System.Text.Encoding.GetEncoding("utf-8").GetString(Convert.FromBase64String(user));
            String rpassword = System.Text.Encoding.GetEncoding("utf-8").GetString(Convert.FromBase64String(password));
            if (textBox2.Text == rpassword)
            {
                MessageBox.Show("登入成功，歡迎回來");
                unlockb = true;
            }
            else
            {
                MessageBox.Show("帳號或密碼錯誤");
                unlockb = false;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Register f4 = new Register();
            f4.Name = "更改密碼";
            f4.ShowDialog();
        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            AboutCoder f6 = new AboutCoder();
            f6.Show();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (button13.Text == "顯示密碼")
            {
                textBox2.UseSystemPasswordChar = false;
                button13.Text = "隱藏密碼";
            }
            else
            {
                textBox2.UseSystemPasswordChar = true;
                button13.Text = "顯示密碼";
            }
        }

        private void label51_Click(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (Glabel != null && Glabel.Tag.ToString() != "null")
                 Glabel.BackColor = DefaultBackColor;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (Glabel != null && Glabel.Tag.ToString() != "null")
                Glabel.BackColor = Color.MediumAquamarine;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            Init();
            if (File.Exists(wpath))
            {
                string[] text = System.IO.File.ReadAllLines(wpath);
                for (int i = 0; i < text.Length; i++)
                {
                    String[] split = text[i].Split(',');
                    int h = Convert.ToInt16(split[0]);
                    int d = Convert.ToInt16(split[1]);
                    String name = split[2].ToString();
                    arraysH[h, d].Text = name.Split(' ')[0] + "\n" + name.Split(' ')[1];
                    arraysH[h, d].Tag = String.Format("{0},{1}", h + 1, d + 1);
                    arraysH[h, d].BackColor = Color.LightGreen;
                }
                label58.Text = "SeatInit";
                for (int i = 0; i < 7; i++)
                {
                    for (int j = 0; j < 7; j++)
                    {
                        if (arraysH[i, j].Text == null || arraysH[i, j].Text == " ")
                        {
                            arraysH[i, j].Tag = "null";
                            arraysH[i, j].BackColor = Color.Gray;
                        }
                    }
                }
            }
        }
        private void button17_Click(object sender, EventArgs e)
        {
            Form8 f8 = new Form8();
            f8.Show();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            ChatSystem f9 = new ChatSystem();
            f9.Show();
        }
    }
}
