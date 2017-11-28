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

namespace WindowsFormsApp1
{
    public partial class NameTagUpload : Form
    {
        public NameTagUpload()
        {
            InitializeComponent();
        }

        public Label[,] arraysH;
        public String lpath,wpath = System.IO.Directory.GetCurrentDirectory() + @"\SeatList.txt",excelpath;
        String[] Flex = new String[49];
        Label label;

        private void Form2_Load(object sender, EventArgs e)
        {
            ArrayInit();
            for (int i = 0; i < 7; i++)
            {
                for (int j = 0; j < 7; j++)
                {
                    arraysH[i, j].Click += new EventHandler(labelClick);
                    arraysH[i, j].Tag = String.Format("{0},{1},{2}", i + 1, j + 1, "空白");
                    arraysH[i, j].Text = null;
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
        List<String> remove = new List<String>();
        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i < comboBox1.Items.Count; i++)
                {
                    comboBox1.Items.Clear();
                }
                string[] lines = System.IO.File.ReadAllLines(lpath);
                bool r = true;
                for (int i = 0; i < lines.Length; i++)
                {
                    if (!lines[i].Contains(" "))
                    {
                        MessageBox.Show("檔案錯誤: 檔案第" + i + 1 + "行沒有空白字元(座號和名稱未分開)");
                        label61.Text = "檔案第" + i + 1 + "行沒有空白字元";
                        r = false;
                        break;
                    }
                }
                if (r == true)
                {
                    for (int i = 0; i < lines.Length; i++)
                    {
                        comboBox1.Items.Insert(i, lines[i]);
                    }
                    for (int i = 0; i < remove.Count; i++)
                    {
                        comboBox1.Items.Remove(remove.ElementAt(i));
                    }
                    label61.Text = "名條已成功更新";
                    MessageBox.Show("名條已成功更新","UpdateSuccessful",MessageBoxButtons.OK,MessageBoxIcon.Information);
                }
            }
            catch (System.UnauthorizedAccessException)
            {
                MessageBox.Show("請檢查您的檔案路徑是否有誤\n\n或者你是否有權限可以開啟此檔案", "錯誤編號001",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            catch (System.IO.FileNotFoundException)
            {
                MessageBox.Show("請檢查您的檔案路徑是否有誤\n\n或者你是否有權限可以開啟此檔案", "錯誤編號001", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch(System.ArgumentNullException)
            {
                MessageBox.Show("請檢查您的檔案路徑是否有誤\n\n或者你是否有權限可以開啟此檔案", "錯誤編號001", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        Label Clabel,Glabel;

        private void button5_Click(object sender, EventArgs e)
        {
            if(comboBox1.Text != "" && comboBox1.Text != null)
            {
                if (Glabel.Text == "" || Glabel.Text == null)
                {
                    String comb = comboBox1.Text;
                    Glabel.Text = comb;
                    String removed = comboBox1.Text.Split(' ')[1];
                    comboBox1.Items.Remove(comb);
                    comboBox1.Text = "";
                    remove.Add(comb);
                }
                else
                {
                    MessageBox.Show("此格已有座位資料，請清除","Error:AlreadyExist",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if(Glabel.Text != null && Glabel.Text != "")
            {
                String s = Glabel.Text;
                int x = Convert.ToInt16(s.Split(' ')[0]);
                comboBox1.Items.Add(s);
                Glabel.Text = "";
                remove.Remove(s);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (File.Exists(wpath))
            {
                File.Delete(wpath);
            }
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(wpath))
            { 
                for(int i = 0; i < 7; i++)
                {
                    for(int j = 0; j < 7; j++)
                    {
                        if (arraysH[i,j].Text != null && arraysH[i,j].Text != "") 
                        file.WriteLine(i + "," + j + "," + arraysH[i,j].Text);
                    }
                }
            }
            CreateNameTag();
            
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.ShowDialog();
            label52.Text = ofd.InitialDirectory + ofd.FileName;
            lpath = ofd.InitialDirectory + ofd.FileName;
        }

        private void label52_Click(object sender, EventArgs e)
        {
            
        }

        private void labelClick(object sender, EventArgs e)
        {
            label = (Label) sender;
            if (label != Clabel && Clabel != null && Clabel.BackColor == Color.Yellow)
            {
                Clabel.BackColor = DefaultBackColor;
            }
            Glabel = label;
            label58.Text = String.Format("第{0}排第{1}個", Convert.ToInt16(label.Tag.ToString().Split(',')[0]), Convert.ToInt16(label.Tag.ToString().Split(',')[1]));
            if (label.BackColor != Color.Empty)
            {
                label.BackColor = Color.Yellow;
                Clabel = label;
            }
        }

        /*Excel.Application excelapp;
        Excel._Workbook wBook;
        Excel._Worksheet wSheet;
        Excel.Range wRange;

        public void ExcelCreate(String path)
        {
            excelapp = new Excel.Application();
            excelapp.Visible = true;
            excelapp.DisplayAlerts = false;
            excelapp.Workbooks.Add(Type.Missing);
            wBook = excelapp.Workbooks[1];
            wBook.Activate();

            string[] lines = System.IO.File.ReadAllLines(lpath);
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date = DateTime.Now;
            Calendar cal = dfi.Calendar;

                wSheet = (Excel._Worksheet)wBook.Worksheets[1];
                wSheet.Name = DateTime.Now.ToString("week" + cal.GetWeekOfYear(date, dfi.CalendarWeekRule, dfi.FirstDayOfWeek));
                wSheet.Activate();
                excelapp.Cells[1, 1] = "test";

                excelapp.Cells[1, 1] = "名稱";
                //1,2 = 星期一，1,3 = 星期二 ...
                for (int i = 2; i < lines.Length + 2; i++)
                {
                    excelapp.Cells[i, 1] = lines.ElementAt(i-2);
                }
                try
                {
                    wBook.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Console.WriteLine("儲存文件於 " + Environment.NewLine + path);
                }
                catch
                {
                    MessageBox.Show("儲存檔案出錯，檔案可能正在使用");
                }
            
            
            wBook.Close(false, Type.Missing, Type.Missing);
            excelapp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelapp = null;
            GC.Collect();
            Console.Read();
        }*/
        public int dayofweek(String s)
        {
            switch (s)
            {
                case "星期一":
                    return 2;
                case "星期二":
                    return 3;
                case "星期三":
                    return 4;
                case "星期四":
                    return 5;
                case "星期五":
                    return 6;
                case "星期六":
                    return 7;
                case "星期日":
                    return 8;
                default:
                    return 9;
            }
        }

        public void CreateNameTag()
        {
            String path = System.IO.Directory.GetCurrentDirectory() + @"\NameTag.txt";
            if(File.Exists(path)){
                File.Delete(path);
            }
            File.Copy(lpath, path);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            NameTagExample f7 = new NameTagExample();
            f7.Show();
        }
    }
}