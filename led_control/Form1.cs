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
using System.IO.Ports;
using System.Runtime.InteropServices;


namespace led_control
{

    public partial class Form1 : Form
    {
        private SerialPort ComDevice = new SerialPort();
        string iniFilePath = "./Set.ini";
        string portName = "COM1";
        string[] timeArray = new string[26];
        System.Timers.Timer aTimer = new System.Timers.Timer();
        int key_count = 23;

        public Form1()
        {
            InitializeComponent();
            InitialConfig();
        }


        private void InitialConfig()
        {
            if (File.Exists(iniFilePath))
            {
                portName = IniFiles.ReadIniData("Port", "端口号", "", iniFilePath);
                Console.WriteLine(portName);
                for (int i = 0; i < key_count; i++)
                {
                    timeArray[i] = IniFiles.ReadIniData("Params", "视频" + i.ToString() + "时间长度单位秒", "", iniFilePath);
                    Console.WriteLine(timeArray[i]);
                }
                ComDevice.PortName = portName;
                ComDevice.BaudRate = 115200;
                ComDevice.Parity = Parity.None;
                ComDevice.DataBits = 8;
                ComDevice.StopBits = StopBits.One;
                try
                {
                    ComDevice.Open();
                }
                catch (Exception ex)
                {
                    log_label.Text = "请检查配置参数，并重新打开！";
                    MessageBox.Show(ex.Message, "请检查配置参数，并重新打开", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ComDevice.DataReceived += new SerialDataReceivedEventHandler(Com_DataReceived);
                log_label.Text = "正在运行！";

                char[] buffer = new char[1];
                buffer[0] = 'e';
                ComDevice.Write(buffer, 0, 1);

                aTimer.Elapsed += new System.Timers.ElapsedEventHandler(TimerHanlder);
            }
            else
            {
                log_label.Text = "无配置文件，已创建配置文件，请配置参数并重新打开！";
                IniFiles.WriteIniData("Port", "端口号", "COM1", iniFilePath);
                for (int i = 0; i < key_count; i++)
                {
                    IniFiles.WriteIniData("Params", "视频" + i.ToString() + "时间长度单位秒", "1", iniFilePath);
                }
            }

        }

        private void TimerHanlder(object source, System.Timers.ElapsedEventArgs e)
        {
            char[] buffer = new char[1];
            buffer[0] = 'e';
            aTimer.Enabled = false;
            ComDevice.Write(buffer, 0, 1);
        }

        private void SetTimerParam(int interval)
        {
            aTimer.Interval = interval * 1000;
            aTimer.AutoReset = false;
            aTimer.Enabled = true;
        }

        private void Com_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            int i = 0;
            byte[] ReDatas = new byte[ComDevice.BytesToRead];
            ComDevice.Read(ReDatas, 0, ReDatas.Length);
            for (i = 0; i < ReDatas.Length; i++)
            {
                switch (ReDatas[i])
                {
                    case (byte)'a':
                        SendKeys.SendWait("{a}");
                        SetTimerParam(Convert.ToInt32(timeArray[0]));
                        break;
                    case (byte)'b':
                        SendKeys.SendWait("{b}");
                        SetTimerParam(Convert.ToInt32(timeArray[1]));
                        break;
                    case (byte)'c':
                        SendKeys.SendWait("{c}");
                        SetTimerParam(Convert.ToInt32(timeArray[2]));
                        break;
                    case (byte)'d':
                        SendKeys.SendWait("{d}");
                        SetTimerParam(Convert.ToInt32(timeArray[3]));
                        break;
                    case (byte)'e':
                        SendKeys.SendWait("{e}");
                        SetTimerParam(Convert.ToInt32(timeArray[4]));
                        break;
                    case (byte)'f':
                        SendKeys.SendWait("{f}");
                        SetTimerParam(Convert.ToInt32(timeArray[5]));
                        break;
                    case (byte)'g':
                        SendKeys.SendWait("{g}");
                        SetTimerParam(Convert.ToInt32(timeArray[6]));
                        break;
                    case (byte)'h':
                        SendKeys.SendWait("{h}");
                        SetTimerParam(Convert.ToInt32(timeArray[7]));
                        break;
                    case (byte)'i':
                        SendKeys.SendWait("{i}");
                        SetTimerParam(Convert.ToInt32(timeArray[8]));
                        break;
                    case (byte)'j':
                        SendKeys.SendWait("{j}");
                        SetTimerParam(Convert.ToInt32(timeArray[9]));
                        break;
                    case (byte)'k':
                        SendKeys.SendWait("{k}");
                        SetTimerParam(Convert.ToInt32(timeArray[10]));
                        break;
                    case (byte)'l':
                        SendKeys.SendWait("{l}");
                        SetTimerParam(Convert.ToInt32(timeArray[11]));
                        break;
                    case (byte)'m':
                        SendKeys.SendWait("{m}");
                        SetTimerParam(Convert.ToInt32(timeArray[12]));
                        break;
                    case (byte)'n':
                        SendKeys.SendWait("{n}");
                        SetTimerParam(Convert.ToInt32(timeArray[13]));
                        break;
                    case (byte)'o':
                        SendKeys.SendWait("{o}");
                        SetTimerParam(Convert.ToInt32(timeArray[14]));
                        break;
                    case (byte)'p':
                        SendKeys.SendWait("{p}");
                        SetTimerParam(Convert.ToInt32(timeArray[15]));
                        break;
                    case (byte)'q':
                        SendKeys.SendWait("{q}");
                        SetTimerParam(Convert.ToInt32(timeArray[16]));
                        break;
                    case (byte)'r':
                        SendKeys.SendWait("{r}");
                        SetTimerParam(Convert.ToInt32(timeArray[17]));
                        break;
                    case (byte)'s':
                        SendKeys.SendWait("{s}");
                        SetTimerParam(Convert.ToInt32(timeArray[18]));
                        break;
                    case (byte)'t':
                        SendKeys.SendWait("{t}");
                        SetTimerParam(Convert.ToInt32(timeArray[19]));
                        break;
                    case (byte)'u':
                        SendKeys.SendWait("{u}");
                        SetTimerParam(Convert.ToInt32(timeArray[20]));
                        break;
                    case (byte)'v':
                        SendKeys.SendWait("{v}");
                        SetTimerParam(Convert.ToInt32(timeArray[21]));
                        break;
                    case (byte)'w':
                        SendKeys.SendWait("{w}");
                        SetTimerParam(Convert.ToInt32(timeArray[22]));
                        break;
                        //case (byte)'x':
                        //    SendKeys.SendWait("{x}");
                        //    SetTimerParam(Convert.ToInt32(timeArray[23]));
                        //    break;
                        //case (byte)'y':
                        //    SendKeys.SendWait("{y}");
                        //    SetTimerParam(Convert.ToInt32(timeArray[24]));
                        //    break;
                        //case (byte)'z':
                        //    SendKeys.SendWait("{z}");
                        //    SetTimerParam(Convert.ToInt32(timeArray[25]));
                        //    break;
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }

    public class IniFiles
    {
        #region API函数声明

        [DllImport("kernel32")]//返回0表示失败，非0为成功
        private static extern long WritePrivateProfileString(string section, string key,
            string val, string filePath);

        [DllImport("kernel32")]//返回取得字符串缓冲区的长度
        private static extern long GetPrivateProfileString(string section, string key,
            string def, StringBuilder retVal, int size, string filePath);


        #endregion

        #region 读Ini文件

        public static string ReadIniData(string Section, string Key, string NoText, string iniFilePath)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, Key, NoText, temp, 1024, iniFilePath);
            return temp.ToString();
        }

        #endregion

        #region 写Ini文件

        public static bool WriteIniData(string Section, string Key, string Value, string iniFilePath)
        {
            long OpStation = WritePrivateProfileString(Section, Key, Value, iniFilePath);
            if (OpStation == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        #endregion
    }
}
