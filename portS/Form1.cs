using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;//用于启用线程类；
using System.IO.Ports;//用于调用串口类函数
using System.Runtime.InteropServices;
using System.IO;
using System.Xml;

namespace portS
{
    
    public partial class Form1 : Form
    {
        int a1, a2, ak1, ak2, b1, b2, bk1, bk2, c1, c2, ck1, ck2, d1, d2, dk1, dk2, e1, e2, ek1, ek2, f1, f2, fk1, fk2;
        bool flag = false;
        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
        public string iPort = "com6"; //默认为串口1
        public int iRate = 9600; //波特率1200,2400,4800,9600
        public byte bSize = 8; //8 bits
        public int iTimeout = 1000; //延时时长
        public SerialPort serialPort1 = new SerialPort();//定义一个串口类的串口变量
        string serialReadString; //用于串口接收数据
        public Thread Thd_Send; //开辟一个专用于发送数据的线程        
        public byte[] recb;  //用于存放接收数据的数组

        public Form1()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;

            XmlDocument doc = new XmlDocument();
            doc.Load(@"ruler.xml");
            XmlNode xn = doc.SelectSingleNode("parameters");
            XmlNodeList xnl = xn.ChildNodes;
            a1 = int.Parse(xnl[0].InnerText);
            a2 = int.Parse(xnl[1].InnerText);
            ak1 = int.Parse(xnl[2].InnerText);
            ak2 = int.Parse(xnl[3].InnerText);

            b1 = int.Parse(xnl[4].InnerText);
            b2 = int.Parse(xnl[5].InnerText);
            bk1 = int.Parse(xnl[6].InnerText);
            bk2 = int.Parse(xnl[7].InnerText);

            c1 = int.Parse(xnl[8].InnerText);
            c2 = int.Parse(xnl[9].InnerText);
            ck1 = int.Parse(xnl[10].InnerText);
            ck2 = int.Parse(xnl[11].InnerText);

            d1 = int.Parse(xnl[12].InnerText);
            d2 = int.Parse(xnl[13].InnerText);
            dk1 = int.Parse(xnl[14].InnerText);
            dk2 = int.Parse(xnl[15].InnerText);

            e1 = int.Parse(xnl[16].InnerText);
            e2 = int.Parse(xnl[17].InnerText);
            ek1 = int.Parse(xnl[18].InnerText);
            ek2 = int.Parse(xnl[19].InnerText);

            f1 = int.Parse(xnl[20].InnerText);
            f2 = int.Parse(xnl[21].InnerText);
            fk1 = int.Parse(xnl[22].InnerText);
            fk2 = int.Parse(xnl[23].InnerText);
            if (a1 < 2)
                a1 = 2;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            //if (!serialPort1.IsOpen)
                //flag = false;
            //if (flag == false)
            {
                Microsoft.VisualBasic.Devices.Computer cmbCOM = new Microsoft.VisualBasic.Devices.Computer();
                System.Collections.ObjectModel.ReadOnlyCollection<string> ports = cmbCOM.Ports.SerialPortNames;
                //int num = ports.Count;
                foreach (string s in cmbCOM.Ports.SerialPortNames)
                {
                    Parity myParity = Parity.None;
                    StopBits MyStopBits = StopBits.One;
                    if (serialPort1.IsOpen)
                    {
                        serialPort1.Close();
                    }

                    serialPort1.PortName = Convert.ToString(s);  //1,2,3,4

                    serialPort1.BaudRate = 115200; //1200,2400,4800,9600
                    int bita = 8;
                    serialPort1.DataBits = Convert.ToByte(bita.ToString(), 10); //8 bits
                    serialPort1.Parity = myParity; // 0-4=no,odd,even,mark,space
                    serialPort1.StopBits = MyStopBits;
                    this.OpenCom();

                    /*if (this.OpenCom())
                        SubSendData();
                    //Thread.Sleep(1000);
                    if (flag)
                    {
                        //timer1.Interval = 300000;
                        break;
                    }*/
                }
            }
        }

        

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            serialPort1.Close();
        }

        int count = 0;
        string flags = "";
        private void OnDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            serialReadString = "";
            serialReadString += serialPort1.ReadExisting();
            string[] sArray = serialReadString.Split('\r');
            if(sArray.Length>1)
                textBox1.Text = sArray[1].Replace("\n", "");
            //textBox1.Text = serialReadString;
            int temp = 0;
            try
            {
                temp = int.Parse(textBox1.Text);
            }
            catch
            { }
            if (temp > a1 && temp < a2)
            {
                if (flags == "a")
                {
                    if (count < 5)
                    {
                        keybd_event((byte)Keys.A, 0, 0, 0);
                        keybd_event((byte)Keys.A, 0, 2, 0);
                        count++;
                    }
                }
                else
                {
                    keybd_event((byte)Keys.A, 0, 0, 0);
                    keybd_event((byte)Keys.A, 0, 2, 0);
                    count = 1;
                    flags = "a";
                }
            }
            else if (temp > b1 && temp < b2)
            {
                if (flags == "b")
                {
                    if (count < 5)
                    {
                        keybd_event((byte)Keys.B, 0, 0, 0);
                        keybd_event((byte)Keys.B, 0, 2, 0);
                        count++;
                    }
                }
                else
                {
                    keybd_event((byte)Keys.B, 0, 0, 0);
                    keybd_event((byte)Keys.B, 0, 2, 0);
                    count = 1;
                    flags = "b";
                }

            }
            else if (temp > c1 && temp < c2)
            {
                if (flags == "c")
                {
                    if (count < 5)
                    {
                        keybd_event((byte)Keys.C, 0, 0, 0);
                        keybd_event((byte)Keys.C, 0, 2, 0);
                        count++;
                    }
                }
                else
                {
                    keybd_event((byte)Keys.C, 0, 0, 0);
                    keybd_event((byte)Keys.C, 0, 2, 0);
                    count = 1;
                    flags = "c";
                }

            }
            else if (temp > d1 && temp < d2)
            {
                if (flags == "d")
                {
                    if (count < 5)
                    {
                        keybd_event((byte)Keys.D, 0, 0, 0);
                        keybd_event((byte)Keys.D, 0, 2, 0);
                        count++;
                    }
                }
                else
                {
                    keybd_event((byte)Keys.D, 0, 0, 0);
                    keybd_event((byte)Keys.D, 0, 2, 0);
                    count = 1;
                    flags = "d";
                }
            }
            else if (temp > e1 && temp < e2)
            {
                if (flags == "e")
                {
                    if (count < 5)
                    {
                        keybd_event((byte)Keys.E, 0, 0, 0);
                        keybd_event((byte)Keys.E, 0, 2, 0);
                        count++;
                    }
                }
                else
                {
                    keybd_event((byte)Keys.E, 0, 0, 0);
                    keybd_event((byte)Keys.E, 0, 2, 0);
                    count = 1;
                    flags = "e";
                }
            }
            else if (temp > f1 && temp < f2)
            {
                if (flags == "f")
                {
                    if (count < 5)
                    {
                        keybd_event((byte)Keys.F, 0, 0, 0);
                        keybd_event((byte)Keys.F, 0, 2, 0);
                        count++;
                    }
                }
                else
                {
                    keybd_event((byte)Keys.F, 0, 0, 0);
                    keybd_event((byte)Keys.F, 0, 2, 0);
                    count = 1;
                    flags = "f";
                }
            }
            else if ((temp > ak1 && temp < ak2) ||
                (temp > bk1 && temp <= bk2) ||
                (temp > ck1 && temp <= ck2) ||
                (temp > dk1 && temp <= dk2) ||
                (temp > ek1 && temp <= ek2) ||
                (temp > fk1 && temp <= fk2))
            {
                if (flags == "k")
                {
                    if (count < 5000)
                    {
                        keybd_event((byte)Keys.K, 0, 0, 0);
                        keybd_event((byte)Keys.K, 0, 2, 0);
                        count++;
                    }
                }
                else
                {
                    keybd_event((byte)Keys.K, 0, 0, 0);
                    keybd_event((byte)Keys.K, 0, 2, 0);
                    count = 1;
                    flags = "k";
                }
            }
        }
        #region 发送

        //发送数据子函数
        public void SubSendData()
        {
            int SendNumB = 0;
            try
            {
                byte[] b = { 0x30, 0x0d, 0x0a};

                serialPort1.Write(b,0,3);
                
                SendNumB = serialPort1.BytesToWrite;
                //Thread.Sleep(1000);
            }
            catch
            {
                return;
            }
        }
        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public bool OpenCom()
        {
            try
            {
                if (!(serialPort1.IsOpen))
                {
                    serialPort1.Open();//打开串口
                }
                return true;
            }
            catch (Exception e)
            {
                //MessageBox.Show("错误：" + e.Message);
                return false;
            }
        }

        //serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(this.OnDataReceived);
    }
}
