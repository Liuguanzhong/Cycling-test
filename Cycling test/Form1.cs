using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace Cycling_test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            serialPort1.Encoding = Encoding.GetEncoding("GB2312");
        }

        private void Delay(int Millisecond) //延迟系统时间，但系统又能同时能执行其它任务；
        {
            DateTime current = DateTime.Now;
            while (current.AddMilliseconds(Millisecond) > DateTime.Now)
            {

                System.Windows.Forms.Application.DoEvents();//转让控制权            
            }
            return;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region [加载全景显示表]
            this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[0].Cells[0].Value = "1号";
            this.dataGridView1.Rows[0].Cells[1].Value = "2号";
            this.dataGridView1.Rows[0].Cells[2].Value = "3号";
            this.dataGridView1.Rows[0].Cells[3].Value = "4号";
            this.dataGridView1.Rows[0].Cells[4].Value = "5号";
            this.dataGridView1.Rows[0].Cells[5].Value = "6号";
            this.dataGridView1.Rows[0].Cells[6].Value = "7号";
            this.dataGridView1.Rows[0].Cells[7].Value = "8号";
            this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[1].Cells[0].Value = "9号";
            this.dataGridView1.Rows[1].Cells[1].Value = "10号";
            this.dataGridView1.Rows[1].Cells[2].Value = "11号";
            this.dataGridView1.Rows[1].Cells[3].Value = "12号";
            this.dataGridView1.Rows[1].Cells[4].Value = "13号";
            this.dataGridView1.Rows[1].Cells[5].Value = "14号";
            this.dataGridView1.Rows[1].Cells[6].Value = "15号";
            this.dataGridView1.Rows[1].Cells[7].Value = "16号";
            this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[2].Cells[0].Value = "17号";
            this.dataGridView1.Rows[2].Cells[1].Value = "18号";
            this.dataGridView1.Rows[2].Cells[2].Value = "19号";
            this.dataGridView1.Rows[2].Cells[3].Value = "20号";
            this.dataGridView1.Rows[2].Cells[4].Value = "21号";
            this.dataGridView1.Rows[2].Cells[5].Value = "22号";
            this.dataGridView1.Rows[2].Cells[6].Value = "23号";
            this.dataGridView1.Rows[2].Cells[7].Value = "24号";
            this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[3].Cells[0].Value = "25号";
            this.dataGridView1.Rows[3].Cells[1].Value = "26号";
            this.dataGridView1.Rows[3].Cells[2].Value = "27号";
            this.dataGridView1.Rows[3].Cells[3].Value = "28号";
            this.dataGridView1.Rows[3].Cells[4].Value = "29号";
            this.dataGridView1.Rows[3].Cells[5].Value = "30号";
            this.dataGridView1.Rows[3].Cells[6].Value = "31号";
            this.dataGridView1.Rows[3].Cells[7].Value = "32号";
            this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[4].Cells[0].Value = "33号";
            this.dataGridView1.Rows[4].Cells[1].Value = "34号";
            this.dataGridView1.Rows[4].Cells[2].Value = "35号";
            this.dataGridView1.Rows[4].Cells[3].Value = "36号";
            this.dataGridView1.Rows[4].Cells[4].Value = "37号";
            this.dataGridView1.Rows[4].Cells[5].Value = "38号";
            this.dataGridView1.Rows[4].Cells[6].Value = "39号";
            this.dataGridView1.Rows[4].Cells[7].Value = "40号";
            this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[5].Cells[0].Value = "41号";
            this.dataGridView1.Rows[5].Cells[1].Value = "42号";
            this.dataGridView1.Rows[5].Cells[2].Value = "43号";
            this.dataGridView1.Rows[5].Cells[3].Value = "44号";
            this.dataGridView1.Rows[5].Cells[4].Value = "45号";
            this.dataGridView1.Rows[5].Cells[5].Value = "46号";
            this.dataGridView1.Rows[5].Cells[6].Value = "47号";
            this.dataGridView1.Rows[5].Cells[7].Value = "48号";
            this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[6].Cells[0].Value = "49号";
            this.dataGridView1.Rows[6].Cells[1].Value = "50号";
            this.dataGridView1.Rows[6].Cells[2].Value = "51号";
            this.dataGridView1.Rows[6].Cells[3].Value = "52号";
            this.dataGridView1.Rows[6].Cells[4].Value = "53号";
            this.dataGridView1.Rows[6].Cells[5].Value = "54号";
            this.dataGridView1.Rows[6].Cells[6].Value = "55号";
            this.dataGridView1.Rows[6].Cells[7].Value = "56号";
            this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[7].Cells[0].Value = "57号";
            this.dataGridView1.Rows[7].Cells[1].Value = "58号";
            this.dataGridView1.Rows[7].Cells[2].Value = "59号";
            this.dataGridView1.Rows[7].Cells[3].Value = "60号";
            this.dataGridView1.Rows[7].Cells[4].Value = "61号";
            this.dataGridView1.Rows[7].Cells[5].Value = "62号";
            this.dataGridView1.Rows[7].Cells[6].Value = "63号";
            this.dataGridView1.Rows[7].Cells[7].Value = "64号";
            this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[8].Cells[0].Value = "65号";
            this.dataGridView1.Rows[8].Cells[1].Value = "66号";
            this.dataGridView1.Rows[8].Cells[2].Value = "67号";
            this.dataGridView1.Rows[8].Cells[3].Value = "68号";
            this.dataGridView1.Rows[8].Cells[4].Value = "69号";
            this.dataGridView1.Rows[8].Cells[5].Value = "70号";
            this.dataGridView1.Rows[8].Cells[6].Value = "71号";
            this.dataGridView1.Rows[8].Cells[7].Value = "72号";
            this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[9].Cells[0].Value = "73号";
            this.dataGridView1.Rows[9].Cells[1].Value = "74号";
            this.dataGridView1.Rows[9].Cells[2].Value = "75号";
            this.dataGridView1.Rows[9].Cells[3].Value = "76号";
            this.dataGridView1.Rows[9].Cells[4].Value = "77号";
            this.dataGridView1.Rows[9].Cells[5].Value = "78号";
            this.dataGridView1.Rows[9].Cells[6].Value = "79号";
            this.dataGridView1.Rows[9].Cells[7].Value = "80号";
            #endregion
            dataGridView1.ClearSelection();
            
            textBox1.AppendText("未连接");
        }

        int linkdmmflag = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            if (linkdmmflag == 0)
            {
                textBox1.Clear();
                SearchAndLinkVC8246BAuto(serialPort1, ackID);
            }
            else if (linkdmmflag == 1)
            {
                button1.Enabled = false;
                button2.Enabled = true;
                button2.BackColor = Color.Red;
                flag_buttoncolor = 1;
                flag_datashow = 1;
                Thread t1 = new Thread(activate_buttoncolor);
                t1.IsBackground = true;
                //Thread t2 = new Thread(datashow_activision);
                t1.Start();//线程开始执行
                //t2.Start();//线程开始执行
                setcyclingtimer();//设置询问万用表的定时器
                build_dataview2();//建立表格2
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            flag_buttoncolor = 0;
            flag_datashow = 0;
            button1.Enabled = true;
            button2.Enabled = false;
            button2.BackColor = Color.Silver;
            button1.BackColor = Color.LawnGreen;
            datashow_activision_clear();
            linkdmmflag = 0;
            serialPort1.Write(rstID, 0, 7);
            serialPort1.Close();
            cyclingTimer.Close();
            row_count = 0;
            column_count = 0;
            //关闭其余的线程
        }

        int flag_buttoncolor = 0;
        private void activate_buttoncolor()
        {
            while (flag_buttoncolor==1)
            {
                Delay(500);
                if (button1.BackColor == Color.LawnGreen)
                {
                    button1.BackColor = Color.Silver;
                }
                else
                {
                    button1.BackColor = Color.LawnGreen;
                }
                
            }
            button1.BackColor = Color.LawnGreen;
        }

        int flag_datashow = 0;
        int row_count = 0;
        int column_count = 0;
        private void datashow_activision()
        {
            
            if (column_count < 8)
            {
                dataGridView1.Rows[row_count].Cells[column_count].Style.BackColor = Color.LawnGreen;
                column_count++;
            }
            else
            {
                column_count = 0;
                row_count++;
                if (row_count>9)
                {
                    row_count = 0;
                    datashow_activision_clear();
                }
                dataGridView1.Rows[row_count].Cells[column_count].Style.BackColor = Color.LawnGreen;
                column_count++;
            }
        }

        private void datashow_activision_clear()
        {
            int row_count = 0;
            int column_count = 0;
            for (row_count = 0; row_count < 10; row_count++)
            {
                for (column_count = 0; column_count < 8; column_count++)
                {
                    dataGridView1.Rows[row_count].Cells[column_count].Style.BackColor = Color.White;
                }
            }
        }

        System.Timers.Timer cyclingTimer = new System.Timers.Timer(1000);//设置间隔时间为1000毫秒
        private void setcyclingtimer()
        {
            //System.Timers.Timer cyclingTimer = new System.Timers.Timer(1000);//设置间隔时间为1000毫秒；
            cyclingTimer.Elapsed += new System.Timers.ElapsedEventHandler(serialport_cycling_send);
            cyclingTimer.AutoReset = true;
            cyclingTimer.Enabled = true;

            cyclingTimer.Start();
        }

        byte[] readadc = { 0x23, 0x2A, 0x52, 0x44, 0x3F, 0x0D, 0x0A };
        int test_number=0;
        int serialport_count=0;
        int cycling_count=0;
        int tempcount=0;
        private void serialport_cycling_send(object source, System.Timers.ElapsedEventArgs e)
        {
            string sret = "";
            if (test_number != 0)
            {
                try
                {
                    lock (serialPort1)
                    {
                        sret = _DialogToDMM(rdID, 1000);
                    }
                    SetText(sret + "\r\n");
                    //这里开线程读串口数据，且要lock，花时间等待返回值
                }
                catch { }

                if (test_number == serialport_count)
                {
                    datashow_activision();
                }

                if ((test_number == serialport_count) && (cycling_count == tempcount))
                {
                    dt2.Rows.Add();
                    SetDGVSourceFunction(dt2);
                    string s1 = (cycling_count + 1).ToString();
                    dt2.Rows[cycling_count][0] = "第" + s1 + "次";
                    SetDGVSourceFunction(dt2);
                    cycling_count++;
                    //tempcount++;
                    return;
                }
            }

            if ((test_number == 0) && (cycling_count == tempcount))
            {
                dt2.Rows.Add();
                SetDGVSourceFunction(dt2);
                dt2.Rows[0][0] = "第1次";
                SetDGVSourceFunction(dt2);
                cycling_count++;
                test_number = 1;
                serialport_count = 1;
                //tempcount++;
                return;
            }

            if (test_number < 20)
            {
                test_number++;
                serialport_count++;
                dt2.Rows[cycling_count-1][test_number-1] = sret;
                SetDGVSourceFunction(dt2);
            }
            else
            {
                tempcount++;
                test_number = 1;
                serialport_count = 1;
            }
        }

        private string _DialogToDMM(byte[] str_string, int nWaitTime = 0, Thread thd = null)
        {
            string sret = "";

            try
            {
                while (serialPort1.BytesToRead > 0) { serialPort1.ReadByte(); }

                //输出SCPI语句
                //serialPort1.WriteLine(scpi_string);

                serialPort1.Write(str_string, 0, 7);
                //  (字节数 / (波特率/11)) * 1000 = 发送完语句需要的毫秒数
                int slen = str_string.Length + 4;
                int bps = serialPort1.BaudRate;
                int stime = slen * 1000 * 11 / bps;
                Thread.Sleep(stime);

                if (thd != null) { thd.Start(); }

                //  等待接收数据
                if (nWaitTime != 0)
                {
                    serialPort1.ReadTimeout = nWaitTime;
                    string sr = serialPort1.ReadLine().Replace("\r", string.Empty);
                    sret = sr;
                }
            }
            catch (Exception ex)
            {
                sret = "DMM ERROR: " + ex.Message;
                //if (_com.IsOpen == false)
                //{
                //    Disconnect();
                //}
            }

            return sret;
        }

        #region [跨线程访问datagridview1]
        DataTable dt1 = new DataTable();
        private delegate void SetDGVSource111(DataTable dt);
        public /*static*/ void SetDGVSourceFunction111(DataTable dt)
        {
            if (dataGridView1.InvokeRequired)
            {
                SetDGVSource111 delegateSetSource = new SetDGVSource111(SetDGVSourceFunction111);
                dataGridView1.BeginInvoke(delegateSetSource, new object[] { dt });
            }
            else
            {
                dataGridView1.DataSource = dt;
                dataGridView1.Refresh();
            }
        }
        #endregion

        #region [跨线程访问datagridview2]
        DataTable dt2 = new DataTable();
        private delegate void SetDGVSource(DataTable dt);//添加设置DataGridView的DataSource的代理
        public /*static*/ void SetDGVSourceFunction(DataTable dt)
        {
            if (dataGridView2.InvokeRequired)
            {
                SetDGVSource delegateSetSource = new SetDGVSource(SetDGVSourceFunction);
                dataGridView2.BeginInvoke(delegateSetSource, new object[] { dt });
            }
            else
            {
                dataGridView2.DataSource = null;//神奇的步骤，解决垂直滚动条失效的问题
                dataGridView2.DataSource = dt;
                dataGridView2.Refresh();
                if (dataGridView2.Rows.Count>3)
                {
                    dataGridView2.FirstDisplayedScrollingRowIndex = dataGridView2.Rows.Count - 1;
                }
            }
        }
        #endregion

        #region[跨线程访问textbox1]
        delegate void SetTextCallback(string text);
        public void SetText(string text)
        {
            if (this.textBox1.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.textBox1.Text = text;
            }
        }
        #endregion

        private void build_dataview2()
        {
            dataGridView2.ReadOnly = true;      //禁用编辑功能
            //dataGridView2.ColumnCount = 21;
            //dataGridView2.ColumnHeadersVisible = false;

            // Set the column header style.
            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

            columnHeaderStyle.BackColor = Color.Beige;
            columnHeaderStyle.Font = new Font("Verdana", 12, FontStyle.Bold);
            dataGridView2.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

            // Set the column header names.
            //dataGridView2.Columns[0].Name = "测试序号";
            dt2.Columns.Add("循环序号", typeof(String));
            SetDGVSourceFunction(dt2);
            for (int i=0;i<20;i++)
            {
                string num = Convert.ToString(i+1);
                string productnum = "产品" + num + "#";
                //dataGridView2.Columns[i+1].Name = productnum;
                dt2.Columns.Add(productnum, typeof(String));
                SetDGVSourceFunction(dt2);
            }
            
        }

        class data_excel
        {
            private string data_adc;
            private int test_number;
            private int cycling_number;
            public data_excel(string Dataadc,int Testnum,int Cyclingnum)
            {
                this.data_adc = Dataadc;
                this.test_number = Testnum;
                this.cycling_number = Cyclingnum;
            }
            public string Dataadc
            {
                get { return data_adc; }
            }
            public int Testnum
            {
                get { return test_number; }
            }
            public int Cyclingnum
            {
                get { return cycling_number; }
            }
        }//构造表格成员
        private List<data_excel> datatoexcel = new List<data_excel>(1024);
        private void Record_to_sheet(string datastring,float datafloat)
        {
            for (int temp=0;temp<1024;temp++)//此处不应有循环
            {
                data_excel tempstring = new data_excel(datastring,temp,temp);
                datatoexcel.Add(tempstring);
            }
        }

        //string datastring;
        //int buffcount = 0;
        //public byte[] usefulbuf;
        //private List<byte> buffer = new List<byte>(4096);//默认分配1页内存，并始终限制不允许超过
        //private void Serialport_Datareceived(object sender, SerialDataReceivedEventArgs e)//接收存储串口数据
        //{
        //    //try
        //    //{
        //        int counts1 = serialPort1.BytesToRead;
        //        buffcount += counts1;
        //        byte[] buffread = new byte[counts1];
        //        serialPort1.Read(buffread, 0, counts1);
        //        buffer.AddRange(buffread);
        //    if (buffcount>4000)
        //    {
        //        buffer.RemoveAt(0);
        //    }
        //        //if (counts1>4)
        //        //    {
        //        //    usefulbuf = new byte[counts1 - 4];
        //        //    serialPort1.Read(buffread, 0, counts1);
        //        //    buffer.AddRange(buffread);
        //        //    }
                
        //        //while (buffer.Count>=5)
        //        //{
        //        //if (buffer[0] == 0x23 && buffer[1] == 0x2a && buffer[2] == 0x52 && buffer[4] == 0x44)
        //        //    {
        //        //        int n = 2;
        //        //        for (n=2;n< counts1; n++)
        //        //        {
        //        //            if (buffer[n]==0x0d)
        //        //            {
        //        //                if (buffer[n + 1] == 0x0a)
        //        //                {
        //        //                    int i;
        //        //                    for (i=0;i< counts1-4; i++)
        //        //                    {
        //        //                        usefulbuf[i] = buffer[i + 4];
        //        //                    }
        //        //                    datastring = System.Text.Encoding.ASCII.GetString(usefulbuf);
        //        //                    test_number = 0;
        //        //                    dt2.Rows[0][1] = datastring;
        //        //                    SetDGVSourceFunction(dt2);
        //        //                    //SetText(datastring+"\r\n");
        //        //                    //float dataread;
        //        //                    //if (!float.TryParse(datastring,out dataread))
        //        //                    //{
        //        //                    //    //数据转换异常
        //        //                    //    break;
        //        //                    //}

        //        //                    //this.Invoke(Handle_data);
        //        //                    buffer.RemoveRange(0,n+1);
        //        //                    return;
        //        //                }
        //        //                else {  }
        //        //            }
        //        //        }
        //        //    }
        //        //    else
        //        //    {
        //        //        buffer.RemoveAt(0);
        //        //        //Thread.Sleep(2);
        //        //    }
        //        //}
        //    //}
        //    //catch { }
        //}

       

        byte[] ackID = { 0x23, 0x2a, 0x06, 0x00, 0x0d, 0x0a };
        byte[] insID = { 0x23, 0x2a, 0x49, 0x4e, 0x53, 0x30, 0x30, 0x30, 0x0d, 0x0a };
        byte[] rdID = { 0x23, 0x2a, 0x52, 0x44, 0x3f, 0x0d, 0x0a };
        byte[] rstID = { 0x23, 0x2a, 0x52, 0x53, 0x54, 0x0d, 0x0a };

        private void SearchAndLinkVC8246BAuto(SerialPort MyPort, byte[] LoadID)//搜索数字万用表
        {
            string Buffer;
            byte[] chipID = { 0, 0, 0, 0, 0, 0 };
            byte[] AskloadID = { 0x23, 0x2a, 0x4f, 0x4e, 0x4c, 0x0d, 0x0a };
            int flag;
            for (int i = 1; i < 99; i++)
            {
                try
                {
                    flag = 0;
                    Buffer = "COM" + i.ToString();
                    MyPort.PortName = Buffer;
                    MyPort.Open();
                    MyPort.DiscardInBuffer();
                    MyPort.ReadTimeout = 100;
                    MyPort.Write(AskloadID, 0, 7);
                    try
                    {
                        Delay(500);
                        MyPort.Read(chipID, 0, 6);
                        if (flag == 0)
                        {
                            if ((chipID[0] == LoadID[0]) && (chipID[1] == LoadID[1]))
                            {
                                MyPort.DiscardInBuffer();
                                chipID = null;
                                if (MyPort == serialPort1)
                                {
                                    MyPort.Write(insID, 0, 10);
                                    linkdmmflag = 1;
                                    textBox1.AppendText("\r\n已连接");
                                }
                                break;
                            }
                            else
                            {
                                MyPort.DiscardInBuffer();
                                MyPort.Close();
                                flag = 1;
                            }
                        }
                    }
                    catch
                    {
                        if (MyPort == serialPort1)
                        { textBox1.AppendText("\r\n通信超时"); }
                        MyPort.DiscardInBuffer();
                        MyPort.Close();
                    }
                }
                catch
                {
                }
                if (i >= 98)
                {
                    //if (MyPort == serialPort2)
                    { textBox1.AppendText("未搜索到对应串口"); }
                }
            }
        }

    }
}
