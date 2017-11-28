using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.IO;
namespace WindowsFormsApp1
{

    public partial class ChatSystem : Form
    {
        Socket socket = null;
        List<String> list = null; 
        public ChatSystem()
        {
            InitializeComponent();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            
            
            String ip = null;
            try
            {
                if (button1.Text == "Connect")
                {
                    socket = new Socket(SocketType.Stream, ProtocolType.Tcp);
                    list = textBox1.Lines.ToList();
                    ip = textBox2.Text;
                    int port = Convert.ToInt16(textBox5.Text);
                    list.Add("連接 " + ip + " 中...");
                    textBox1.Lines = list.ToArray();
                    IPAddress ipa = Dns.GetHostAddresses(ip)[0];
                    socket.Connect(ipa, port);
                    list.Add("連接成功!");
                    textBox1.Lines = list.ToArray();
                    button1.Text = "Disconnect";
                    ThreadMain tm = new ThreadMain(socket);
                    Thread thread = new Thread(new ThreadStart(tm.run));
                    thread.Start();
                }
                else
                {
                    socket.Disconnect(true);
                    socket.Close();
                    socket.Dispose();
                    list.Add("連接結束，已停止。");
                    textBox1.Lines = list.ToArray();
                    button1.Text = "Connect";
                }
            }
            catch (System.Net.Sockets.SocketException)
            {
                list.Add("無法連接到" + ip);
                textBox1.Lines = list.ToArray();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string Utext = File.ReadAllText(System.IO.Directory.GetCurrentDirectory() + @"\site.txt");
            String user = Utext.Split(',')[0];
            String name = System.Text.Encoding.GetEncoding("utf-8").GetString(Convert.FromBase64String(user));
            byte[] b = new byte[1024];
            byte[] text = Encoding.UTF8.GetBytes("[" + name + "] " + textBox3.Text);
            socket.Send(text);
            list.Add("[" + name + "] " + textBox3.Text);
            textBox1.Lines = list.ToArray();
        }
    }
    public class ThreadMain
    {
        Socket s;
        Thread t;
        public ThreadMain(Socket s)
        {
            this.s = s;
            ThreadSecond ts = new ThreadSecond(s);
            t = new Thread(new ThreadStart(ts.run));
        }
        public void run()
        {
            while (true)
            {
                t.Start();
            }
        }
    }
    public class ThreadSecond
    {
        Socket s;
        public ThreadSecond(Socket s)
        {
            this.s = s;
        }
        public void run()
        {
            while (true)
            {
                byte[] b = new byte[1024];
                int text = s.Receive(b);
                String str = Encoding.UTF8.GetString(b, 0, text);
                
            }
        }
    }
}
