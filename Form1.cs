using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Net.Sockets;
using System.Data;
using System.Drawing;
using System.Text;
using System.Diagnostics.Metrics;
using System.Text.Json.Nodes;
using System.Diagnostics.Eventing.Reader;
using System.Threading.Tasks;
using System.Threading;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using System.Windows.Forms;
using Advantech.Adam;              // 用於Advantech ADAM模組
using Advantech.Common;       // 用於Advantech通用功能
using Advantech.Graph;
using Advantech.Protocol;
using cszmcaux;                             //用於 ECI0064IO控制器 模組功能
using CsvHelper;
using Microsoft.VisualBasic.Logging;
using Microsoft.Identity.Client.NativeInterop;
using System.Threading.Tasks.Dataflow;
using System.Data.Entity.Core.Metadata.Edm;
using System.Windows.Markup;


/***************************************************************************
//連線 ECI0064 及TABLE 說明

//string ipaddr ="192.168.0.11";  //ECI0064 實際IO控制器  
//string phandle="ECI0064C" ;     //權柄設定

//連線ADAM-6017
//private string szip1 ="192.168.0.12"; //ADAM-6017E  1~6 類比電壓  7~8類比電流   第五組
//private string szip2 ="192.168.0.13"; //ADAM-6017A 1~6 類比電壓  7~8類比電流   第一組
//private string szip3 ="192.168.0.14"; //ADAM-6017B 1~6 類比電壓  7~8類比電流   第二組
//private string szip4 ="192.168.0.15"; //ADAM-6017C 1~6 類比電壓  7~8類比電流   第三組
//private string szip5 ="192.168.0.16"; //ADAM-6017D 1~6 類比電壓  7~8類比電流   第四組
//private int m_DeviceFwVer = 4;
//private int m_iPort;
//private int m_iCount;
//private int m_iAiTotal, m_iDoTotal;
//private bool[] m_bChEnabled;
//private byte[] m_byRange;
//private ushort[] m_usRange; //for newer version
****************************************************************************/
namespace ECI0064
{
    public partial class Form1 : Form
    {
        /// <summary>//ADAM-6017

        /// </summary>//ADAM-6017
        /*設定CSV 檔案路徑*/
        private DataTable DataTableCsv = new DataTable();
        private string csvPathFile = @"C:\Work\Chelicdb.csv";
        //private List<string[]> workLog = new List<string[]>();
        //dataTableCsv name
        //判斷是有變更測試品名稱是否有變更
        private bool isDataChanged = false;
        //判斷是有變更測試品名稱是否有變更
        string Device = "";
        string Aps = "";
        string WorkStatus = "";
        string WorkDT = "";
        string WorkNum = "0";
        string WorkUnm = "0";
        //
        /* 啟動設定變數*/
        string ipaddr;                                                          //ECI0064 模擬器IP 位置
        string Schang = "a";                                             //抓夾起始設定
        string workstr;                                                       //日期狀態

        /*全作動 變數*/
        int TCount = 1000;                                              // TCount 保壓時間 TCount *1000= Thread.Sleep(TCount)  Thread.Sleep(1000)=1S ;
        int Amun = 10;                                                      //保壓次數
        int Bmun = 20;                                                      //動作次數
        int Wtime = 1;                                                       //工作時間  Wtime*1000= Thread.Sleep(Wtime)       Thread.Sleep(1000)=1S ;
        int Totalaction;                                                    // A測試品 累計作動次數 Totalaction= ( j-1)*Bmun + i;
        int ACount = 100;                                               //動作次數;
        int AWtime = 1;                                                    //自作動時間  AWtimea*100= Thread.Sleep(WtimeA)   Thread.Sleep(100)=10ms ;
        bool passAll = false;                                            // 全作動判斷是否暫停
        /*全作動變數暫存值*/
        int TempTotalaction;                                        // A測試品 累計作動次數 Totalaction= ( j-1)*Bmun + i;
        /*A測試品 變數*/
        int TCountA = 1000;                                           // TCounA 保壓時間 TCountA *1000= Thread.Sleep(TCountA)  Thread.Sleep(1000)=1S ;
        int AmunA;                                                          //保壓次數
        int BmunA;                                                           //動作次數
        int WtimeA = 1;                                                    //工作時間  WtimeA*1000= Thread.Sleep(WtimeA)       Thread.Sleep(1000)=1S ;
        int Totalactiona;                                                   // A測試品 累計作動次數 Totalaction= ( ja-1)*BmunA + ia;
        int ACounta = 100;                                             //動作次數;
        int AWtimea = 1;                                                  //自作動時間  AWtimea*100= Thread.Sleep(WtimeA)   Thread.Sleep(100)=10ms ;    
        string DeviceA1;                                                 //測試品A1 對應 TextA1  暫存 TextA1Data 
        string DeviceA2;                                                 //測試品A2 對應 TextA2  暫存 TextA2Data 
        string DeviceA3;                                                 //測試品A3 對應 TextA3  暫存 TextA3Data 
        bool passA = false;                                              //判斷是否A測試品 暫停
        /*A測試品暫存變數*/
        int TempTotalactiona;                                          // A測試品 累計作動次數 Totalaction= ( ja-1)*BmunA + ia;      
        /*B測試品 變數*/
        int TCountB = 1000;                                           // TCountB 保壓時間 TCountB *1000= Thread.Sleep(TCountB)  Thread.Sleep(1000)=1S ;
        int AmunB;                                                           //保壓次數
        int BmunB;                                                           //動作次數
        int WtimeB = 1;                                                    //工作時間  Wtime*1000= Thread.Sleep(Wtime)       Thread.Sleep(1000)=1S ;
        int Totalactionb;                                                 // B測試品 累計作動次數 Totalaction= ( jb-1)*BmunB + ib;
        int ACountb = 100;                                            //動作次數;
        int AWtimeb = 1;                                                 //自作動時間  AWtimeb*100= Thread.Sleep(WtimeB)   Thread.Sleep(100)=10ms ;
        string DeviceB1;                                                 //測試品B1 對應 TextB1  暫存 TextB1Data 
        string DeviceB2;                                                 //測試品B2 對應 TextB2 暫存 TextB2Data 
        string DeviceB3;                                                 //測試品B3 對應 TextB3  暫存 TextB3Data 
        bool passB = false;                                              //判斷是否B測試品 暫停      
        /*B測試品暫存變數*/
        int TempTotalactionb;                                          // A測試品 累計作動次數 Totalaction= ( ja-1)*BmunA + ia;   ; 
        /*C測試品 變數*/
        int TCountC = 1000;                                         // TCountC 保壓時間 TCountC *1000= Thread.Sleep(TCountC)  Thread.Sleep(1000)=1S ;
        int AmunC;                                                           //保壓次數
        int BmunC;                                                           //動作次數
        int WtimeC = 1;                                                  //工作時間  Wtime*1000= Thread.Sleep(Wtime)       Thread.Sleep(1000)=1S ;
        int Totalactionc;                                                // C測試品 累計作動次數 Totalaction= ( jc-1)*BmunC + ic;
        int ACountc = 100;                                           //動作次數;
        int AWtimec = 1;                                                //自作動時間  AWtimec*100= Thread.Sleep(WtimeC)   Thread.Sleep(100)=10ms ;
        string DeviceC1;                                                 //測試品C1 對應 TextC1  暫存 TextC1Data 
        string DeviceC2;                                                 //測試品C2 對應 TextC2  暫存 TextC2Data 
        string DeviceC3;                                                 //測試品C3 對應 TextC3  暫存 TextC3Data 
        bool passC = false;                                           //判斷是否C測試品 暫停     
        /*C測試品暫存變數*/
        int TempTotalactionc;                                            // A測試品 累計作動次數 Totalaction= ( ja-1)*BmunA + ia;
        /*D測試品 變數*/
        int TCountD = 1000;                                        // TCountD 保壓時間 TCountD *1000= Thread.Sleep(TCountD)  Thread.Sleep(1000)=1S ;
        int AmunD = 10;                                                //保壓次數
        int BmunD = 20;                                                //動作次數
        int WtimeD = 1;                                                 //工作時間  Wtime*1000= Thread.Sleep(Wtime)       Thread.Sleep(1000)=1S ;
        int Totalactiond;                                               // D測試品 累計作動次數 Totalaction= ( jd-1)*BmunD + id;
        int ACountd = 100;                                          //動作次數;
        int AWtimed = 1;                                               //自作動時間  AWtimed*100= Thread.Sleep(WtimeD)   Thread.Sleep(100)=10ms ;
        string DeviceD1;                                                 //測試品D1 對應 TextD1  暫存 TextD1Data 
        string DeviceD2;                                                 //測試品D2 對應 TextD2  暫存 TextD2Data
        string DeviceD3;                                                 //測試品D3 對應 TextD3  暫存 TextD3Data
        bool passD = false;                                            //判斷是否D測試品 暫停
        /*D測試品暫存變數*/
        int TempTotalactiond;                                            // A測試品 累計作動次數 Totalaction= ( ja-1)*BmunA + ia;  
        /*E測試品 變數*/
        int TCountE = 1000;                                         // TCountE 保壓時間 TCountE *1000= Thread.Sleep(TCountE)  Thread.Sleep(1000)=1S ;
        int AmunE = 10;                                                 //保壓次數
        int BmunE = 20;                                                 //動作次數
        int WtimeE = 1;                                                  //工作時間  Wtime*1000= Thread.Sleep(Wtime)       Thread.Sleep(1000)=1S ;
        int Totalactione;                                               //  E測試品  累計作動次數 Totalaction= ( je-1)*BmunE + ie;
        int ACounte = 100;                                          //動作次數;
        int AWtimee = 1;                                               //自作動時間  AWtimee*100= Thread.Sleep(WtimeE)   Thread.Sleep(100)=10ms ;
        string DeviceE1;                                                 //測試品E1 對應 TextE1  暫存 TextE1Data
        string DeviceE2;                                                 //測試品E2 對應 TextE2  暫存 TextE2Data
        string DeviceE3;                                                 //測試品E3 對應 TextE3  暫存 TextE3Data
        bool passE = false;                                            //判斷是否E測試品 暫停      
        /*E測試品暫存變數*/
        int TempTotalactione;                                            // A測試品 累計作動次數 Totalaction= ( ja-1)*BmunA + ia;   

        /*資燉庫資訊*/
        //    string ConStr = "Provider=SQLOLEDB.1;Persist Security Info=True;Intital Catalog=CHELICTM;Data Source=192.168.0.10";
        //    string SqlPwd ='';
        //     string Sqld="Chelctm";
        //     string SqlStr; 
        /*資料庫資料*/
        //Adam6017 設定值
       
        //Adam6017 設定值
        public Form1()
        {
            InitializeComponent();
            SetcCvdata();                    //設定dataGridView DataSource 設定dataTable  
            LoadCsvdata();                //讀入CSV資料
            SorDataTable();               //排序
            dataGridView1.DataSource = DataTableCsv;
           //取壓力表的值
            Adam6017a();
            Adam6017b();
            Adam6017c();
            Adam6017d();
            Adam6017e();
        }

        //ADAM-6017A
        private void Adam6017a()
        {                   
            // 創建一個 Adam6000 類型的對象，用於與 ADAM-6017 模組通信
            AdamSocket adamModbus = new AdamSocket(AdamType.Adam6000);
            // 設置模組的 IP 地址和端口
            string ip = "192.168.0.13"; // 替換為您的 ADAM-6017 模組的 IP 地址
            int port = 502; // Modbus TCP 通常使用的端口
            AnalogInput analogin = new AnalogInput(adamModbus);
            if (adamModbus.Connect(ip, ProtocolType.Tcp, port))
            {
                Console.WriteLine("連接成功！");
                
                byte[] buffer = new byte[16];   // 8個通道，每個通道2個字節
                double[] Napdata = new double[8];
                // 讀取模組的模擬輸入值 資料
                if (adamModbus.Modbus().ReadInputRegs(1,8,out buffer))
                {
                    Console.WriteLine("讀取資料成功！");
                   
                    for (int i = 0; i < buffer.Length/2 ;  i++)
                    {
                       ushort rawValue=BitConverter.ToUInt16(buffer, i*2);
                       double voltage = rawValue * (5-0.5)/ 65535.0+0.5 ;
                        Napdata[i] = voltage*1.82;
                        Console.WriteLine($"Ch- {i}  : {voltage}");                        
                    }
                                Napa1.Text = Napdata[0].ToString("0.##");                               
                                Napa2.Text = Napdata[1].ToString("0.##");
                                Napa3.Text = Napdata[2].ToString("0.##");
                                Napa4.Text = Napdata[3].ToString("0.##");                 
                }
                else
                { Console.WriteLine("讀取資料失敗！"); }
                // 斷開連接
                adamModbus.Disconnect();
            }
            else
            {
                Console.WriteLine("連接失敗！");
            }
           
            
            return;
        }
        //ADAM-6017A
        //ADAM-6017B
        private void Adam6017b()
        {
            // 創建一個 Adam6000 類型的對象，用於與 ADAM-6017 模組通信
            AdamSocket adamModbus = new AdamSocket(AdamType.Adam6000);
            // 設置模組的 IP 地址和端口
            string ip = "192.168.0.14"; // 替換為您的 ADAM-6017 模組的 IP 地址
            int port = 502; // Modbus TCP 通常使用的端口
            AnalogInput analogin = new AnalogInput(adamModbus);
            if (adamModbus.Connect(ip, ProtocolType.Tcp, port))
            {
                Console.WriteLine("連接成功！");

                byte[] buffer = new byte[16];   // 8個通道，每個通道2個字節
                double[] Napdata = new double[8];
                // 讀取模組的模擬輸入值 資料
                if (adamModbus.Modbus().ReadInputRegs(1, 8, out buffer))
                {
                    Console.WriteLine("讀取資料成功！");

                    for (int i = 0; i < buffer.Length / 2; i++)
                    {
                        ushort rawValue = BitConverter.ToUInt16(buffer, i * 2);
                        double voltage = rawValue * (5 - 0.5) / 65535.0 + 0.5;
                        Napdata[i] = voltage * 1.82;
                        Console.WriteLine($"Ch- {i}  : {voltage}");
                    }
                    Napb1.Text = Napdata[0].ToString("0.##");
                    Napb2.Text = Napdata[1].ToString("0.##");
                    Napb3.Text = Napdata[2].ToString("0.##");
                    Napb4.Text = Napdata[3].ToString("0.##");
                }
                else
                { Console.WriteLine("讀取資料失敗！"); }
                // 斷開連接
                adamModbus.Disconnect();
            }
            else
            {
                Console.WriteLine("連接失敗！");
            }
             return;
        }
        //ADAM-6017B
        //ADAM-6017C
        private void Adam6017c()
        {
            // 創建一個 Adam6000 類型的對象，用於與 ADAM-6017 模組通信
            AdamSocket adamModbus = new AdamSocket(AdamType.Adam6000);
            // 設置模組的 IP 地址和端口
            string ip = "192.168.0.15"; // 替換為您的 ADAM-6017 模組的 IP 地址
            int port = 502; // Modbus TCP 通常使用的端口
            AnalogInput analogin = new AnalogInput(adamModbus);
            if (adamModbus.Connect(ip, ProtocolType.Tcp, port))
            {
                Console.WriteLine("連接成功！");

                byte[] buffer = new byte[16];   // 8個通道，每個通道2個字節
                double[] Napdata = new double[8];
                // 讀取模組的模擬輸入值 資料
                if (adamModbus.Modbus().ReadInputRegs(1, 8, out buffer))
                {
                    Console.WriteLine("讀取資料成功！");

                    for (int i = 0; i < buffer.Length / 2; i++)
                    {
                        ushort rawValue = BitConverter.ToUInt16(buffer, i * 2);
                        double voltage = rawValue * (5 - 0.5) / 65535.0 + 0.5;
                        Napdata[i] = voltage * 1.82;
                        Console.WriteLine($"Ch- {i}  : {voltage}");
                    }
                    Napc1.Text = Napdata[0].ToString("0.##");
                    Napc2.Text = Napdata[1].ToString("0.##");
                    Napc3.Text = Napdata[2].ToString("0.##");
                    Napc4.Text = Napdata[3].ToString("0.##");
                }
                else
                { Console.WriteLine("讀取資料失敗！"); }
                // 斷開連接
                adamModbus.Disconnect();
            }
            else
            {
                Console.WriteLine("連接失敗！");
            }
            return;
        }
        //ADAM-6017C
        //ADAM-6017D
        private void Adam6017d()
        {
            // 創建一個 Adam6000 類型的對象，用於與 ADAM-6017 模組通信
            AdamSocket adamModbus = new AdamSocket(AdamType.Adam6000);
            // 設置模組的 IP 地址和端口
            string ip = "192.168.0.16"; // 替換為您的 ADAM-6017 模組的 IP 地址
            int port = 502; // Modbus TCP 通常使用的端口
            AnalogInput analogin = new AnalogInput(adamModbus);
            if (adamModbus.Connect(ip, ProtocolType.Tcp, port))
            {
                Console.WriteLine("連接成功！");

                byte[] buffer = new byte[16];   // 8個通道，每個通道2個字節
                double[] Napdata = new double[8];
                // 讀取模組的模擬輸入值 資料
                if (adamModbus.Modbus().ReadInputRegs(1, 8, out buffer))
                {
                    Console.WriteLine("讀取資料成功！");

                    for (int i = 0; i < buffer.Length / 2; i++)
                    {
                        ushort rawValue = BitConverter.ToUInt16(buffer, i * 2);
                        double voltage = rawValue * (5 - 0.5) / 65535.0 + 0.5;
                        Napdata[i] = voltage * 1.82;
                        Console.WriteLine($"Ch- {i}  : {voltage}");
                    }
                    Napd1.Text = Napdata[0].ToString("0.##");
                    Napd2.Text = Napdata[1].ToString("0.##");
                    Napd3.Text = Napdata[2].ToString("0.##");
                    Napd4.Text = Napdata[3].ToString("0.##");
                }
                else
                { Console.WriteLine("讀取資料失敗！"); }
                // 斷開連接
                adamModbus.Disconnect();
            }
            else
            {
                Console.WriteLine("連接失敗！");
            }
            return;
        }
        //ADAM-6017D
        //ADAM-6017E
        private void Adam6017e()
        {
            // 創建一個 Adam6000 類型的對象，用於與 ADAM-6017 模組通信
            AdamSocket adamModbus = new AdamSocket(AdamType.Adam6000);
            // 設置模組的 IP 地址和端口
            string ip = "192.168.0.12"; // 替換為您的 ADAM-6017 模組的 IP 地址
            int port = 502; // Modbus TCP 通常使用的端口
            AnalogInput analogin = new AnalogInput(adamModbus);
            if (adamModbus.Connect(ip, ProtocolType.Tcp, port))
            {
                Console.WriteLine("連接成功！");

                byte[] buffer = new byte[16];   // 8個通道，每個通道2個字節
                double[] Napdata = new double[8];
                // 讀取模組的模擬輸入值 資料
                if (adamModbus.Modbus().ReadInputRegs(1, 8, out buffer))
                {
                    Console.WriteLine("讀取資料成功！");

                    for (int i = 0; i < buffer.Length / 2; i++)
                    {
                        ushort rawValue = BitConverter.ToUInt16(buffer, i * 2);
                        double voltage = rawValue * (5 - 0.5) / 65535.0 + 0.5;
                        Napdata[i] = voltage * 1.82;
                        Console.WriteLine($"Ch- {i}  : {voltage}");
                    }
                    Nape1.Text = Napdata[0].ToString("0.##");
                    Nape2.Text = Napdata[1].ToString("0.##");
                    Nape3.Text = Napdata[2].ToString("0.##");
                    Nape4.Text = Napdata[3].ToString("0.##");
                }
                else
                { Console.WriteLine("讀取資料失敗！"); }
                // 斷開連接
                adamModbus.Disconnect();
            }
            else
            {
                Console.WriteLine("連接失敗！");
            }
            return;
        }
        //ADAM-6017E
        //設定datagridview 標題
        private void SetcCvdata()
        {
            //設定dataGridView DataSource 設定dataTable  
            dataGridView1.Columns.Clear();

            // 創建測試品列並設定 DataPropertyName
            DataGridViewTextBoxColumn deviceColumn = new DataGridViewTextBoxColumn();
            deviceColumn.HeaderText = "測試品";
            deviceColumn.DataPropertyName = "Device";
            dataGridView1.Columns.Add(deviceColumn);

            // 創建壓力列並設定 DataPropertyName
            DataGridViewTextBoxColumn apsColumn = new DataGridViewTextBoxColumn();
            apsColumn.HeaderText = "壓力";
            apsColumn.DataPropertyName = "Aps";
            dataGridView1.Columns.Add(apsColumn);

            // 創建工作狀態列並設定 DataPropertyName
            DataGridViewTextBoxColumn workStatusColumn = new DataGridViewTextBoxColumn();
            workStatusColumn.HeaderText = "工作狀態";
            workStatusColumn.DataPropertyName = "WorkStatus";
            dataGridView1.Columns.Add(workStatusColumn);

            // 創建工作時間列並設定 DataPropertyName
            DataGridViewTextBoxColumn workDTColumn = new DataGridViewTextBoxColumn();
            workDTColumn.HeaderText = "工作時間";
            workDTColumn.DataPropertyName = "WorkDT";
            dataGridView1.Columns.Add(workDTColumn);

            // 創建工作次數列並設定 DataPropertyName
            DataGridViewTextBoxColumn workNumColumn = new DataGridViewTextBoxColumn();
            workNumColumn.HeaderText = "工作次數";
            workNumColumn.DataPropertyName = "WorkNum";
            dataGridView1.Columns.Add(workNumColumn);

            // 創建保壓次數列並設定 DataPropertyName
            DataGridViewTextBoxColumn workUnmColumn = new DataGridViewTextBoxColumn();
            workUnmColumn.HeaderText = "保壓次數";
            workUnmColumn.DataPropertyName = "WorkUnm";
            dataGridView1.Columns.Add(workUnmColumn);

            dataGridView1.AutoGenerateColumns = false;

            //dataGridView 對應 DataTable 參數
            DataTableCsv.Columns.Add("Device", typeof(string));
            DataTableCsv.Columns.Add("Aps", typeof(string));
            DataTableCsv.Columns.Add("WorkStatus", typeof(string));
            DataTableCsv.Columns.Add("WorkDT", typeof(string));
            DataTableCsv.Columns.Add("WorkNum", typeof(string));
            DataTableCsv.Columns.Add("WorkUnm", typeof(string));

            // 確保 dataGridView1 更新後選擇第一行
            if (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0];
                dataGridView1.Rows[0].Selected = true;
            }

            return;
        }//SetcCvdata()

        //剛開始讀取CSV資料
        private void LoadCsvdata()
        {
            //讀入CSV檔案
            var lines = File.ReadAllLines(csvPathFile);
            if (lines.Length > 0)
            {
                // 使用第一行作為列標題
                string[] columnNames = lines[0].Split(',');
                for (int i = 1; i < lines.Length; i++)
                {
                    string[] rowValues = lines[i].Split(',');
                    DataRow row = DataTableCsv.NewRow();
                    for (int j = 0; j < columnNames.Length; j++)
                    {
                        // 確保不會超出 rowValues 的範圍
                        if (j < rowValues.Length)
                        {
                            row[columnNames[j]] = rowValues[j];
                        }
                        else
                        {
                            // 為缺失的數據添加預設值，例如空字符串
                            row[columnNames[j]] = string.Empty;
                        }
                    }
                    DataTableCsv.Rows.Add(row);
                }
            }
            this.dataGridView1.RowsDefaultCellStyle.BackColor = Color.BurlyWood;                //DataGridView 顏色切換
            this.dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.White;    // 單數行顏色設定
            // 綁定數據源
            dataGridView1.DataSource = DataTableCsv;
            return;
        }//LoadCsvdata()

        private void WireCvsdata()
        {
            DataRow newRow = DataTableCsv.NewRow();
            newRow["Device"] = Device;
            newRow["Aps"] = Aps;
            newRow["WorkStatus"] = WorkStatus;
            newRow["WorkDT"] = WorkDT;
            newRow["WorkNum"] = WorkNum;
            newRow["WorkUnm"] = WorkUnm;
            //
            DataTableCsv.Rows.Add(newRow);
            //指向第一筆資料 
            dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0];
            dataGridView1.Rows[0].Selected = true;
            // 將新的 DataRow 附加到 CSV 檔案中
            using (StreamWriter sw = new StreamWriter(csvPathFile, true)) // true 表示追加模式
            {
                var fields = newRow.ItemArray.Select(field => field.ToString());
                sw.WriteLine(string.Join(",", fields));
            }
            return;
        }

        //資料排序
        private void SorDataTable()
        {
            DataTableCsv.DefaultView.Sort = "WorkDT DESC";// ASC 升序
                                                          ////指向第一筆資料 
            dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0];
            dataGridView1.Rows[0].Selected = true;
            //指向第一筆資料 
            return;
        }

        public void Nowtimed(string workstr)  //取得時間並存取資料
        {
            DateTime currentTime = DateTime.Now;
            string formattedTime = currentTime.ToString("yyyy/MM/dd HH:mm:ss");
            WorkDT = formattedTime;
            Console.WriteLine($"{workstr}時間日期：" + formattedTime);

        }
        // 測試第一組動作  D0 為第一組主電磁閥  D1動作次數
        private async void Bu_tona_Click(object sender, EventArgs e)
        {
            //測試品變數設定            
            Bu_sp.Enabled = false;
            Bu_auto.Enabled = false;
            Bu_allclose.Enabled = false;
            Bu_pas1.Enabled = true;
            ABox.Enabled = false;
            DeviceA1 = TextA1.Text;    //調理
            DeviceA2 = TextA2.Text;    //單口電磁閥
            DeviceA3 = TextA3.Text;    //制動元件
            AmunA = Convert.ToInt32(textBox2.Text);   //保壓次數
            BmunA = Convert.ToInt32(textBox3.Text);   //動作次數
            TCountA = 1000 * Convert.ToInt32(textBox4.Text); //保壓等待時間
            WtimeA = 100 * Convert.ToInt32(Wtime1.Text);    //動作時間
            Awtext.Text = "測試開始";//A測試品      
            workstr = "測試A 開始";
            Adam6017a();       
            WorkNum = Convert.ToString(Totalactiona);
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceA1;       //調理名稱
            Aps = Napa1.Text;         //調理壓力表
            WireCvsdata();
            Device = DeviceA2;
            Aps = Napa2.Text;        //單口電磁關壓力表
            WireCvsdata();
            Device = DeviceA3;
            Aps = Napa3.Text;        //制動元件壓力表
            WireCvsdata();
            SorDataTable();
            Awtext.Refresh();
            /* 重要:加入此段可以讓執行序使用UI */
            Form.CheckForIllegalCrossThreadCalls = false;
            Task<int> A1 = new Task<int>(() => { TONA(); return 0; });
            A1.Start();
            await A1;
            // 測試結束  關閉電磁閥
            if (passA == true)
            {
                Awtext.Text = "測試停止";
                ABox.Enabled = true;
                passA = false;
                Bu_pas1.Enabled = false;
                workstr = "測試A 暫停";
                Adam6017a();             
                Nowtimed(workstr);
                WorkNum = Convert.ToString(TempTotalactiona);
                WorkStatus = workstr;
                Device = DeviceA1;
                Aps = Napa1.Text;         //調理壓力表
                WireCvsdata();
                Device = DeviceA2;
                Aps = Napa2.Text;        //單口電磁關壓力表
                WireCvsdata();
                Device = DeviceA3;
                Aps = Napa3.Text;        //制動元件壓力表
                WireCvsdata();
                SorDataTable();
                Awtext.Refresh();
                Console.WriteLine("測試停止 2 passA =" + passA);//檢查狀態
                if (Bu_tona.Enabled == true && Bu_tonb.Enabled == true && Bu_tonc.Enabled == true && Bu_tond.Enabled == true && Bu_tone.Enabled == true)
                {
                    Bu_sp.Enabled = true;
                    Bu_auto.Enabled = true;
                }
                int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rssult == 0)
                {
                    for (int i = 0; i <= 3; i++)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                    }
                    zmcaux.ZAux_Close(ECI0064C);
                    //  MessageBox.Show("全歸零成功");
                }
                else
                {
                    //       MessageBox.Show("歸零失敗");
                }
                return;
            }
            int rss4 = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C1);
            if (rss4 == 0)
            {
                for (int i = 0; i <= 3; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C1, i, 0);
                }
                zmcaux.ZAux_Close(ECI0064C1);
            }
            Awtext.Text = "測試完成"; // A測試品       
            workstr = "測試A 完成";
            Adam6017a();       
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceA1;
            Aps = Napa1.Text;         //調理壓力表
            WireCvsdata();
            Device = DeviceA2;
            Aps = Napa2.Text;        //單口電磁關壓力表
            WireCvsdata();
            Device = DeviceA3;
            Aps = Napa3.Text;        //制動元件壓力表
            WireCvsdata();
            SorDataTable();
            Awtext.Refresh();
            Totalactiona = 0;
            TempTotalactiona = 0;
            WorkUnm = "";
            WorkNum = "";
            ABox.Enabled = true;
            Bu_pas1.Enabled = false;
            Console.WriteLine("A測試品  測試完成 3 passA = " + passA);//檢查狀態           
            if (Bu_tona.Enabled == true && Bu_tonb.Enabled == true && Bu_tonc.Enabled == true && Bu_tond.Enabled == true && Bu_tone.Enabled == true)
            {
                Bu_sp.Enabled = true;
                Bu_auto.Enabled = true;
            }
            return;
        }  //結束按鈕

        // 測試第二組動作  D4 為第二組主電磁閥  
        private async void Bu_tonb_Click(object sender, EventArgs e)
        {
            Bu_pas2.Enabled = true;
            Bu_sp.Enabled = false;
            Bu_auto.Enabled = false;
            Bu_allclose.Enabled = false;
            BBox.Enabled = false;
            DeviceB1 = TextB1.Text;
            DeviceB2 = TextB2.Text;
            DeviceB3 = TextB3.Text;
            AmunB = Convert.ToInt32(textBox29.Text);   //保壓次數
            BmunB = Convert.ToInt32(textBox28.Text);   //動作次數
            TCountB = 1000 * Convert.ToInt32(textBox30.Text); //保壓等待時間
            WtimeB = 100 * Convert.ToInt32(Wtime2.Text); //動作時間
            Bwtext.Text = "測試開始";//B測試品      
            workstr = "測試B 開始";
             Adam6017b();            
            WorkNum = Convert.ToString(TempTotalactionb);
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceB1;
            Aps = Napb1.Text;         //調理壓力表
            WireCvsdata();
            Device = DeviceB2;
            Aps = Napb2.Text;        //單口電磁關壓力表
            WireCvsdata();
            Device = DeviceB3;
            Aps = Napb3.Text;        //制動元件壓力表
            WireCvsdata();
            SorDataTable();
            Bwtext.Refresh();
            /* 重要:加入此段可以讓執行序使用UI */
            Form.CheckForIllegalCrossThreadCalls = false;
            Task<int> B1 = new Task<int>(() => { TONB(); return 0; });
            B1.Start();
            await B1;
            if (passB == true)
            {
                Bwtext.Text = "測試停止";
                workstr = "測試B 暫停";
                Nowtimed(workstr);
                WorkStatus = workstr;
                WorkNum = Convert.ToString(TempTotalactionb);
                Adam6017b();               
                Device = DeviceB1;
                Aps = Napb1.Text;         //調理壓力表
                WireCvsdata();
                Device = DeviceB2;
                Aps = Napb2.Text;        //單口電磁關壓力表
                WireCvsdata();
                Device = DeviceB3;
                Aps = Napb3.Text;        //制動元件壓力表
                WireCvsdata();
                SorDataTable();
                BBox.Enabled = true;
                passB = false;
                Bu_pas2.Enabled = false;
                if (Bu_tona.Enabled == true && Bu_tonb.Enabled == true && Bu_tonc.Enabled == true && Bu_tond.Enabled == true && Bu_tone.Enabled == true)
                {
                    Bu_sp.Enabled = true;
                    Bu_auto.Enabled = true;
                }
                Console.WriteLine("B測試品 測試暫停  passB = " + passB);//檢查狀態
                int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rssult == 0)
                {
                    for (int i = 4; i <= 7; i++)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                    }
                    zmcaux.ZAux_Close(ECI0064C);
                    // MessageBox.Show("全歸零成功");
                }
                else
                {
                    // MessageBox.Show("歸零失敗");
                }
                return;
            }
            int rss = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C1);
            if (rss == 0)
            {
                for (int i = 4; i <= 7; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C1, i, 0);
                }
                zmcaux.ZAux_Close(ECI0064C1);
            }
            Bwtext.Text = "測試完成"; //B測試品
            workstr = "測試B 完成";          
            Adam6017b();           
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceB1;
            Aps = Napb1.Text;         //調理壓力表
            WireCvsdata();
            Device = DeviceB2;
            Aps = Napb2.Text;        //單口電磁關壓力表
            WireCvsdata();
            Device = DeviceB3;
            Aps = Napb3.Text;        //制動元件壓力表
            WireCvsdata();
            SorDataTable();
            Totalactionb = 0;
            TempTotalactionb = 0;
            WorkUnm = "";
            WorkNum = "";
            BBox.Enabled = true;
            Bu_pas2.Enabled = false;
            if (Bu_tona.Enabled == true && Bu_tonb.Enabled == true && Bu_tonc.Enabled == true && Bu_tond.Enabled == true && Bu_tone.Enabled == true)
            {
                Bu_sp.Enabled = true;
                Bu_auto.Enabled = true;
            }
            Console.WriteLine(" B測試品 測試完成 3 passB = " + passB);//檢查狀態
            return;
        }

        // 測試第三組動作  D8 為第三組主電磁閥    
        private async void Bu_tonc_Click(object sender, EventArgs e)
        {
            Bu_sp.Enabled = false;
            Bu_auto.Enabled = false;
            Bu_pas3.Enabled = true;
            Bu_allclose.Enabled = false;
            CBox.Enabled = false;
            DeviceC1 = TextC1.Text;
            DeviceC2 = TextC2.Text;
            DeviceC3 = TextC3.Text;
            AmunC = Convert.ToInt32(textBox31.Text);   //保壓次數
            BmunC = Convert.ToInt32(textBox32.Text);   //動作次數
            TCountC = 1000 * Convert.ToInt32(textBox33.Text); //保壓等待時間
            WtimeC = 100 * Convert.ToInt32(Wtime3.Text); //保壓等待時間
            Cwtext.Text = "測試開始";//C測試品           
            workstr = "測試C 開始";
            Adam6017c();
            WorkNum = Convert.ToString(TempTotalactionc);
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceC1;
            Aps = Napc1.Text;         //調理壓力表
            WireCvsdata();
            Device = DeviceC2;
            Aps = Napc2.Text;        //單口電磁關壓力表
            WireCvsdata();
            Device = DeviceC3;
            Aps = Napc3.Text;        //制動元件壓力表
            WireCvsdata();
            SorDataTable();
            Cwtext.Refresh();
            /* 重要:加入此段可以讓執行序使用UI */
            Form.CheckForIllegalCrossThreadCalls = false;
            Task<int> C1 = new Task<int>(() => { TONC(); return 0; });
            C1.Start();
            await C1;
            if (passC == true)
            {
                Cwtext.Text = "測試停止";
                workstr = "測試C 暫停";
                Adam6017c();
                Nowtimed(workstr);
                WorkStatus = workstr;
                WorkNum = Convert.ToString(TempTotalactionc);
                Device = DeviceC1;
                Aps = Napc1.Text;         //調理壓力表
                WireCvsdata();
                Device = DeviceC2;
                Aps = Napc2.Text;        //單口電磁關壓力表
                WireCvsdata();
                Device = DeviceC3;
                Aps = Napc3.Text;        //制動元件壓力表
                WireCvsdata();
                SorDataTable();
                CBox.Enabled = true;
                passC = false;
                Bu_pas3.Enabled = false;
                if (Bu_tona.Enabled == true && Bu_tonb.Enabled == true && Bu_tonc.Enabled == true && Bu_tond.Enabled == true && Bu_tone.Enabled == true)
                {
                    Bu_sp.Enabled = true;
                    Bu_auto.Enabled = true;
                }
                Console.WriteLine("C測試品 測試停止 3 passC = " + passC);//檢查狀態
                int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rssult == 0)
                {
                    for (int i = 8; i <= 11; i++)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                    }
                    zmcaux.ZAux_Close(ECI0064C);
                    //  MessageBox.Show("全歸零成功");
                }
                else
                {
                    //       MessageBox.Show("歸零失敗");
                }
                return;
            }
            int rss4 = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C1);
            if (rss4 == 0)
            {
                for (int i = 8; i <= 11; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C1, i, 0);
                }
                zmcaux.ZAux_Close(ECI0064C1);
            }
            Cwtext.Text = "測試完成";//C測試品
            workstr = "測試C 完成";
            Adam6017c();
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceC1;
            Aps = Napc1.Text;         //調理壓力表
            WireCvsdata();
            Device = DeviceC2;
            Aps = Napc2.Text;        //單口電磁關壓力表
            WireCvsdata();
            Device = DeviceC3;
            Aps = Napc3.Text;        //制動元件壓力表
            WireCvsdata();
            SorDataTable();
            Totalactionc = 0;
            TempTotalactionc = 0;
            WorkUnm = "";
            WorkNum = "";
            CBox.Enabled = true;
            Bu_pas3.Enabled = false;
            passC = false;
            if (Bu_tona.Enabled == true && Bu_tonb.Enabled == true && Bu_tonc.Enabled == true && Bu_tond.Enabled == true && Bu_tone.Enabled == true)
            {
                Bu_sp.Enabled = true;
                Bu_auto.Enabled = true;
            }
            Console.WriteLine("C測試品 測試完成 3 passC = " + passC);//檢查狀態
            return;
        }

        private async void Bu_tond_Click(object sender, EventArgs e) // 第四組測試開始
        {
            Bu_pas4.Enabled = true;
            DBox.Enabled = false;
            Bu_sp.Enabled = false;
            Bu_auto.Enabled = false;
            Bu_allclose.Enabled = false;
            DeviceD1 = TextD1.Text;
            DeviceD2 = TextD2.Text;
            DeviceD3 = TextD3.Text; ;
            AmunD = Convert.ToInt32(textBox24.Text);   //保壓次數
            BmunD = Convert.ToInt32(textBox23.Text);   //動作次數
            TCountD = 1000 * Convert.ToInt32(textBox22.Text); //保壓等待時間
            WtimeD = 100 * Convert.ToInt32(Wtime4.Text); //保壓等待時間
            Dwtext.Text = "測試開始";//D測試品
            workstr = "測試D 開始";
            Adam6017d();
            WorkNum = Convert.ToString(TempTotalactiond);
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceD1;
            Aps = Napd1.Text;         //調理壓力表
            WireCvsdata();
            Device = DeviceD2;
            Aps = Napd2.Text;        //單口電磁關壓力表
            WireCvsdata();
            Device = DeviceD3;
            Aps = Napd3.Text;        //制動元件壓力表
            WireCvsdata();
            SorDataTable();
            /* 重要:加入此段可以讓執行序使用UI */
            Form.CheckForIllegalCrossThreadCalls = false;
            Task<int> D1 = new Task<int>(() => { TOND(); return 0; });
            D1.Start();
            await D1;
            if (passD == true)
            {
                Dwtext.Text = " 測試停止";//D測試品
                workstr = "測試D 暫停";
                Adam6017d();
                Nowtimed(workstr);
                WorkStatus = workstr;
                WorkNum = Convert.ToString(TempTotalactiond);
                Device = DeviceD1;
                Aps = Napd1.Text;         //調理壓力表
                WireCvsdata();
                Device = DeviceD2;
                Aps = Napd2.Text;        //單口電磁關壓力表
                WireCvsdata();
                Device = DeviceD3;
                Aps = Napd3.Text;        //制動元件壓力表
                WireCvsdata();
                SorDataTable();
                DBox.Enabled = true;
                passD = false;
                Bu_pas4.Enabled = false;
                if (Bu_tona.Enabled == true && Bu_tonb.Enabled == true && Bu_tonc.Enabled == true && Bu_tond.Enabled == true && Bu_tone.Enabled == true)
                {
                    Bu_sp.Enabled = true;
                    Bu_auto.Enabled = true;
                }
                Console.WriteLine("D測試品 測試停止 3 passD = " + passD);//檢查狀態
                int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rssult == 0)
                {
                    for (int i = 12; i <= 15; i++)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                    }
                    zmcaux.ZAux_Close(ECI0064C);
                    //  MessageBox.Show("全歸零成功");
                }
                else
                {
                    //       MessageBox.Show("歸零失敗");
                }
                return;
            }  //暫停跳出
            int rss4 = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C1);
            if (rss4 == 0)
            {
                for (int i = 12; i <= 15; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C1, i, 0);
                }
                zmcaux.ZAux_Close(ECI0064C1);
            }
            Dwtext.Text = " 測試完成";//D測試品                
            workstr = "測試D 完成";
            Adam6017d();
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceD1;
            Aps = Napd1.Text;         //調理壓力表
            WireCvsdata();
            Device = DeviceD2;
            Aps = Napd2.Text;        //單口電磁關壓力表
            WireCvsdata();
            Device = DeviceD3;
            Aps = Napd3.Text;        //制動元件壓力表
            WireCvsdata();
            SorDataTable();
            Totalactiond = 0;
            TempTotalactiond = 0;
            WorkUnm = "";
            WorkNum = "";
            DBox.Enabled = true;
            Bu_pas4.Enabled = false;
            if (Bu_tona.Enabled == true && Bu_tonb.Enabled == true && Bu_tonc.Enabled == true && Bu_tond.Enabled == true && Bu_tone.Enabled == true)
            {
                Bu_sp.Enabled = true;
                Bu_auto.Enabled = true;
            }
            return;
        }

        private async void Bu_tone_Click(object sender, EventArgs e)//第五組測試開始
        {
            Bu_pas5.Enabled = true;
            EBox.Enabled = false;
            Bu_sp.Enabled = false;
            Bu_auto.Enabled = false;
            Bu_allclose.Enabled = false;
            DeviceE1 = TextE1.Text;
            DeviceE2 = TextE2.Text;
            DeviceE3 = TextE3.Text;
            AmunE = Convert.ToInt32(textBox37.Text);   //保壓次數
            BmunE = Convert.ToInt32(textBox36.Text);   //動作次數
            TCountE = 1000 * Convert.ToInt32(textBox35.Text); //保壓等待時間
            WtimeE = 100 * Convert.ToInt32(Wtime5.Text); //保壓等待時間           
            Ewtext.Text = "測試開始";//E測試品
            workstr = "測試E 開始";
            Adam6017e();
            WorkNum = Convert.ToString(TempTotalactione);
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceE1;
            Aps = Nape1.Text;         //調理壓力表
            WireCvsdata();
            Device = DeviceE2;
            Aps = Nape2.Text;        //單口電磁關壓力表
            WireCvsdata();
            Device = DeviceE3;
            Aps = Nape3.Text;        //制動元件壓力表
            WireCvsdata();
            SorDataTable();
            Ewtext.Refresh();
            /* 重要:加入此段可以讓執行序使用UI */
            Form.CheckForIllegalCrossThreadCalls = false;
            Task<int> E1 = new Task<int>(() => { TONE(); return 0; });
            E1.Start();
            await E1;
            if (passE == true)
            {
                Ewtext.Text = "測試停止";//E測試品
                workstr = "測試E 暫停";
                Adam6017e();
                Nowtimed(workstr);
                WorkStatus = workstr;
                WorkNum = Convert.ToString(TempTotalactione);
                Device = DeviceE1;
                Aps = Nape1.Text;         //調理壓力表
                WireCvsdata();
                Device = DeviceE2;
                Aps = Nape2.Text;        //單口電磁關壓力表
                WireCvsdata();
                Device = DeviceE3;
                Aps = Nape3.Text;        //制動元件壓力表
                WireCvsdata();
                SorDataTable();
                EBox.Enabled = true;
                passE = false;
                Bu_pas5.Enabled = false;
                if (Bu_tona.Enabled == true && Bu_tonb.Enabled == true && Bu_tonc.Enabled == true && Bu_tond.Enabled == true && Bu_tone.Enabled == true)
                {
                    Bu_sp.Enabled = true;
                    Bu_auto.Enabled = true;
                }
                Console.WriteLine("E測試品 測試停止");
                int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rssult == 0)
                {
                    for (int i = 16; i <= 19; i++)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                    }
                    zmcaux.ZAux_Close(ECI0064C);
                    //  MessageBox.Show("全歸零成功");
                }
                else
                {
                    //       MessageBox.Show("歸零失敗");
                }
                return;
            }  //暫停跳出
            int rss4 = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C1);
            if (rss4 == 0)
            {
                for (int i = 16; i <= 19; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C1, i, 0);
                }
                zmcaux.ZAux_Close(ECI0064C1);
            }
            Ewtext.Text = "測試完成";//E測試品  
            workstr = "測試E 完成";
            Adam6017e();
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceE1;
            Aps = Nape1.Text;         //調理壓力表
            WireCvsdata();
            Device = DeviceE2;
            Aps = Nape2.Text;        //單口電磁關壓力表
            WireCvsdata();
            Device = DeviceE3;
            Aps = Nape3.Text;        //制動元件壓力表
            WireCvsdata();
            SorDataTable();
            Totalactione = 0;
            TempTotalactione = 0;
            WorkUnm = "";
            WorkNum = "";
            EBox.Enabled = true;
            Bu_pas5.Enabled = false;
            if (Bu_tona.Enabled == true && Bu_tonb.Enabled == true && Bu_tonc.Enabled == true && Bu_tond.Enabled == true && Bu_tone.Enabled == true)
            {
                Bu_sp.Enabled = true;
                Bu_auto.Enabled = true;
            }
            Console.WriteLine("E測試品 測試完成");
            return;
            //  MessageBox.Show("氣閥測試完成");// 測試結束
        }

        private async void button15_Click(object sender, EventArgs e)
        {
            Bu_allclose.Enabled = true;
            Bu_auto.Enabled = false;
            Bu_sp.Enabled = false;
            ABox.Enabled = false;
            BBox.Enabled = false;
            CBox.Enabled = false;
            DBox.Enabled = false;
            EBox.Enabled = false;
            DeviceA1 = TextA1.Text;
            DeviceA2 = TextA2.Text;
            DeviceA3 = TextA3.Text;
            DeviceB1 = TextB1.Text;
            DeviceB2 = TextB2.Text;
            DeviceB3 = TextB3.Text;
            DeviceC1 = TextC1.Text;
            DeviceC2 = TextC2.Text;
            DeviceC3 = TextC3.Text;
            DeviceD1 = TextD1.Text;
            DeviceD2 = TextD2.Text;
            DeviceD3 = TextD3.Text;
            DeviceE1 = TextE1.Text;
            DeviceE2 = TextE2.Text;
            DeviceE3 = TextE3.Text;
            Awtext.Text = "測試開始"; // A測試品
            Bwtext.Text = "測試開始"; //B測試品
            Cwtext.Text = "測試開始";//C測試品
            Dwtext.Text = "測試開始";//D測試品
            Ewtext.Text = "測試開始";//E測試品
            workstr = "自動開始";
            Adam6017a();
            Adam6017b();
            Adam6017c();
            Adam6017d();
            Adam6017e();
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceA1;
            Aps = Napa1.Text;
            WireCvsdata();
            Device = DeviceA2;
            Aps = Napa2.Text;
            WireCvsdata();
            Device = DeviceA3;
            Aps = Napa3.Text;
            WireCvsdata();
            Device = DeviceB1;
            Aps = Napb1.Text;
            WireCvsdata();
            Device = DeviceB2;
            Aps = Napb2.Text;
            WireCvsdata();
            Device = DeviceB3;
            Aps = Napb3.Text;
            WireCvsdata();
            Device = DeviceC1;
            Aps = Napc1.Text;
            WireCvsdata();
            Device = DeviceC2;
            Aps = Napc2.Text;
            WireCvsdata();
            Device = DeviceC3;
            Aps = Napc3.Text;
            WireCvsdata();
            Device = DeviceD1;
            Aps = Napd1.Text;
            WireCvsdata();
            Device = DeviceD2;
            Aps = Napd2.Text;
            WireCvsdata();
            Device = DeviceD3;
            Aps = Napd3.Text;
            WireCvsdata();
            Device = DeviceE1;
            Aps = Nape1.Text;
            WireCvsdata();
            Device = DeviceE2;
            Aps = Nape2.Text;
            WireCvsdata();
            Device = DeviceE3;
            Aps = Nape3.Text;
            WireCvsdata();
            SorDataTable();
            //判斷是否中斷後在執行
            Amun = Convert.ToInt32(textBox2.Text);   //保壓次數
            Bmun = Convert.ToInt32(textBox3.Text);   //動作次數
            TCount = 1000 * Convert.ToInt32(textBox4.Text); //保壓等待時間
            Wtime = 100 * Convert.ToInt32(Wtime1.Text);    //動作時間
            zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
            Form.CheckForIllegalCrossThreadCalls = false;
            Task<int> All = new Task<int>(() => { TONAll(); return 0; });
            All.Start();
            await All;
            if (passAll == true)
            {
                Awtext.Text = "測試停止"; // A測試品
                Bwtext.Text = "測試停止"; //B測試品
                Cwtext.Text = "測試停止";//C測試品
                Dwtext.Text = "測試停止";//D測試品
                Ewtext.Text = "測試停止";//E測試品
                workstr = "自動 暫停";
                Adam6017a();
                Adam6017b();
                Adam6017c();
                Adam6017d();
                Adam6017e();
                Nowtimed(workstr);
                WorkStatus = workstr;
                Device = DeviceA1;
                Aps = Napa1.Text;
                WireCvsdata();
                Device = DeviceA2;
                Aps = Napa2.Text;
                WireCvsdata();
                Device = DeviceA3;
                Aps = Napa3.Text;
                WireCvsdata();
                Device = DeviceB1;
                Aps = Napb1.Text;
                WireCvsdata();
                Device = DeviceB2;
                Aps = Napb2.Text;
                WireCvsdata();
                Device = DeviceB3;
                Aps = Napb3.Text;
                WireCvsdata();
                Device = DeviceC1;
                Aps = Napc1.Text;
                WireCvsdata();
                Device = DeviceC2;
                Aps = Napc2.Text;
                WireCvsdata();
                Device = DeviceC3;
                Aps = Napc3.Text;
                WireCvsdata();
                Device = DeviceD1;
                Aps = Napd1.Text;
                WireCvsdata();
                Device = DeviceD2;
                Aps = Napd2.Text;
                WireCvsdata();
                Device = DeviceD3;
                Aps = Napd3.Text;
                WireCvsdata();
                Device = DeviceE1;
                Aps = Nape1.Text;
                WireCvsdata();
                Device = DeviceE2;
                Aps = Nape2.Text;
                WireCvsdata();
                Device = DeviceE3;
                Aps = Nape3.Text;
                WireCvsdata();
                SorDataTable();
                ABox.Enabled = true;
                BBox.Enabled = true;
                CBox.Enabled = true;
                DBox.Enabled = true;
                EBox.Enabled = true;
                Bu_sp.Enabled = true;
                Bu_auto.Enabled = true;
                Bu_allclose.Enabled = false;
                passAll = false;
                int rssult = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                if (rssult == 0)
                {
                    for (int i = 0; i <= 31; i++)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                    }
                    zmcaux.ZAux_Close(ECI0064C);
                    //  MessageBox.Show("全歸零成功");
                }
                else
                {
                    //       MessageBox.Show("歸零失敗");
                }
                Console.WriteLine("測試暫停  passALL = " + passAll);//檢查狀態
                return;
            }  //  if (passAll == true) 暫停跳出
            int rss4 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
            if (rss4 == 0)
            {
                for (int i = 0; i <= 31; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                }
                zmcaux.ZAux_Close(ECI0064C);
            }
            Awtext.Text = "測試完成"; // A測試品
            Bwtext.Text = "測試完成"; //B測試品
            Cwtext.Text = "測試完成";//C測試品
            Dwtext.Text = "測試完成";//D測試品
            Ewtext.Text = "測試完成";//E測試品         
            workstr = "自動 完成";
            Adam6017a();
            Adam6017b();
            Adam6017c();
            Adam6017d();
            Adam6017e();
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceA1;
            Aps = Napa1.Text;
            WireCvsdata();
            Device = DeviceA2;
            Aps = Napa2.Text;
            WireCvsdata();
            Device = DeviceA3;
            Aps = Napa3.Text;
            WireCvsdata();
            Device = DeviceB1;
            Aps = Napb1.Text;
            WireCvsdata();
            Device = DeviceB2;
            Aps = Napb2.Text;
            WireCvsdata();
            Device = DeviceB3;
            Aps = Napb3.Text;
            WireCvsdata();
            Device = DeviceC1;
            Aps = Napc1.Text;
            WireCvsdata();
            Device = DeviceC2;
            Aps = Napc2.Text;
            WireCvsdata();
            Device = DeviceC3;
            Aps = Napc3.Text;
            WireCvsdata();
            Device = DeviceD1;
            Aps = Napd1.Text;
            WireCvsdata();
            Device = DeviceD2;
            Aps = Napd2.Text;
            WireCvsdata();
            Device = DeviceD3;
            Aps = Napd3.Text;
            WireCvsdata();
            Device = DeviceE1;
            Aps = Nape1.Text;
            WireCvsdata();
            Device = DeviceE2;
            Aps = Nape2.Text;
            WireCvsdata();
            Device = DeviceE3;
            Aps = Nape3.Text;
            WireCvsdata();
            SorDataTable();
            Totalaction = 0;
            TempTotalaction = 0;
            WorkUnm = "";
            WorkNum = "";
            ABox.Enabled = true;
            BBox.Enabled = true;
            CBox.Enabled = true;
            DBox.Enabled = true;
            EBox.Enabled = true;
            Bu_sp.Enabled = true;
            Bu_auto.Enabled = true;
            Bu_allclose.Enabled = false;
            passAll = false;
            Console.WriteLine("測試完成 passALL =" + passAll);//檢查狀態
        }

        private void button22_Click(object sender, EventArgs e)
        {
            /*pass 為true 所有動作停止*/
            passAll = true;
            Bu_allclose.Enabled = false;
            Bu_auto.Enabled = true;
            workstr = "自動 暫停";
            Adam6017a();
            Adam6017b();
            Adam6017c();
            Adam6017d();
            Adam6017e();
            WorkNum = Convert.ToString(TempTotalaction);
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceA1;
            Aps = Napa1.Text;
            WireCvsdata();
            Device = DeviceA2;
            Aps = Napa2.Text;
            WireCvsdata();
            Device = DeviceA3;
            Aps = Napa3.Text;
            WireCvsdata();
            Device = DeviceB1;
            Aps = Napb1.Text;
            WireCvsdata();
            Device = DeviceB2;
            Aps = Napb2.Text;
            WireCvsdata();
            Device = DeviceB3;
            Aps = Napb3.Text;
            WireCvsdata();
            Device = DeviceC1;
            Aps = Napc1.Text;
            WireCvsdata();
            Device = DeviceC2;
            Aps = Napc2.Text;
            WireCvsdata();
            Device = DeviceC3;
            Aps = Napc3.Text;
            WireCvsdata();
            Device = DeviceD1;
            Aps = Napd1.Text;
            WireCvsdata();
            Device = DeviceD2;
            Aps = Napd2.Text;
            WireCvsdata();
            Device = DeviceD3;
            Aps = Napd3.Text;
            WireCvsdata();
            Device = DeviceE1;
            Aps = Nape1.Text;
            WireCvsdata();
            Device = DeviceE2;
            Aps = Nape2.Text;
            WireCvsdata();
            Device = DeviceE3;
            Aps = Nape3.Text;
            WireCvsdata();
            SorDataTable();
            int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
            if (rssult == 0)
            {
                for (int i = 0; i <= 31; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                }
                zmcaux.ZAux_Close(ECI0064C);
                //  MessageBox.Show("全歸零成功");
            }
            else
            {
                //       MessageBox.Show("歸零失敗");
            }
        }

        private async void button23_Click(object sender, EventArgs e)  // AUTO TEST連動測試
        {
            Bu_allclose.Enabled = true;
            Bu_auto.Enabled = false;
            Bu_sp.Enabled = false;
            ABox.Enabled = false;
            BBox.Enabled = false;
            CBox.Enabled = false;
            DBox.Enabled = false;
            EBox.Enabled = false;
            DeviceA1 = TextA1.Text;
            DeviceA2 = TextA2.Text;
            DeviceA3 = TextA3.Text;
            DeviceB1 = TextB1.Text;
            DeviceB2 = TextB2.Text;
            DeviceB3 = TextB3.Text;
            DeviceC1 = TextC1.Text;
            DeviceC2 = TextC2.Text;
            DeviceC3 = TextC3.Text;
            DeviceD1 = TextD1.Text;
            DeviceD2 = TextD2.Text;
            DeviceD3 = TextD3.Text;
            DeviceE1 = TextE1.Text;
            DeviceE2 = TextE2.Text;
            DeviceE3 = TextE3.Text;
            Awtext.Text = "測試開始"; // A測試品
            Bwtext.Text = "測試開始"; //B測試品
            Cwtext.Text = "測試開始";//C測試品
            Dwtext.Text = "測試開始";//D測試品
            Ewtext.Text = "測試開始";//E測試品         
            workstr = "快速作動 開始";
            Adam6017a();
            Adam6017b();
            Adam6017c();
            Adam6017d();
            Adam6017e();
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceA1;
            Aps = Napa1.Text;
            WireCvsdata();
            Device = DeviceA2;
            Aps = Napa2.Text;
            WireCvsdata();
            Device = DeviceA3;
            Aps = Napa3.Text;
            WireCvsdata();
            Device = DeviceB1;
            Aps = Napb1.Text;
            WireCvsdata();
            Device = DeviceB2;
            Aps = Napb2.Text;
            WireCvsdata();
            Device = DeviceB3;
            Aps = Napb3.Text;
            WireCvsdata();
            Device = DeviceC1;
            Aps = Napc1.Text;
            WireCvsdata();
            Device = DeviceC2;
            Aps = Napc2.Text;
            WireCvsdata();
            Device = DeviceC3;
            Aps = Napc3.Text;
            WireCvsdata();
            Device = DeviceD1;
            Aps = Napd1.Text;
            WireCvsdata();
            Device = DeviceD2;
            Aps = Napd2.Text;
            WireCvsdata();
            Device = DeviceD3;
            Aps = Napd3.Text;
            WireCvsdata();
            Device = DeviceE1;
            Aps = Nape1.Text;
            WireCvsdata();
            Device = DeviceE2;
            Aps = Nape2.Text;
            WireCvsdata();
            Device = DeviceE3;
            Aps = Nape3.Text;
            WireCvsdata();
            SorDataTable();
            ACount = Convert.ToInt32(textBox3.Text);    //動作次數;
            AWtime = Convert.ToInt32(Wtime1.Text);    //自作動時間  AWtimea*100= Thread.Sleep(WtimeA)   Thread.Sleep(100)=10ms ;
            Form.CheckForIllegalCrossThreadCalls = false;
            Task<int> AutoD = new Task<int>(() => { AutoStar(); return 0; });
            AutoD.Start();
            await AutoD;
            if (passAll == true)
            {
                Awtext.Text = "測試停止"; // A測試品
                Bwtext.Text = "測試停止"; //B測試品
                Cwtext.Text = "測試停止";//C測試品
                Dwtext.Text = "測試停止";//D測試品
                Ewtext.Text = "測試停止";//E測試品
                workstr = "快速作動暫停";
                Adam6017a();
                Adam6017b();
                Adam6017c();
                Adam6017d();
                Adam6017e();
                WorkNum = Convert.ToString(TempTotalaction);
                Nowtimed(workstr);
                WorkStatus = workstr;
                Device = DeviceA1;
                Aps = Napa1.Text;
                WireCvsdata();
                Device = DeviceA2;
                Aps = Napa2.Text;
                WireCvsdata();
                Device = DeviceA3;
                Aps = Napa3.Text;
                WireCvsdata();
                Device = DeviceB1;
                Aps = Napb1.Text;
                WireCvsdata();
                Device = DeviceB2;
                Aps = Napb2.Text;
                WireCvsdata();
                Device = DeviceB3;
                Aps = Napb3.Text;
                WireCvsdata();
                Device = DeviceC1;
                Aps = Napc1.Text;
                WireCvsdata();
                Device = DeviceC2;
                Aps = Napc2.Text;
                WireCvsdata();
                Device = DeviceC3;
                Aps = Napc3.Text;
                WireCvsdata();
                Device = DeviceD1;
                Aps = Napd1.Text;
                WireCvsdata();
                Device = DeviceD2;
                Aps = Napd2.Text;
                WireCvsdata();
                Device = DeviceD3;
                Aps = Napd3.Text;
                WireCvsdata();
                Device = DeviceE1;
                Aps = Nape1.Text;
                WireCvsdata();
                Device = DeviceE2;
                Aps = Nape2.Text;
                WireCvsdata();
                Device = DeviceE3;
                Aps = Nape3.Text;
                WireCvsdata();
                SorDataTable();
                ABox.Enabled = true;
                BBox.Enabled = true;
                CBox.Enabled = true;
                DBox.Enabled = true;
                EBox.Enabled = true;
                Bu_sp.Enabled = true;
                passAll = false;
                Bu_auto.Enabled = true;
                Bu_allclose.Enabled = false;
                int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rssult == 0)
                {
                    for (int i = 0; i <= 31; i++)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                    }
                    zmcaux.ZAux_Close(ECI0064C);
                    //  MessageBox.Show("全歸零成功");
                }
                else
                {
                    //       MessageBox.Show("歸零失敗");
                }
                Console.WriteLine("測試暫停 passALL : " + passAll);    //檢查狀態
                return;
            }  //暫停跳出
            Awtext.Text = "連動測試完成 " + ACount + " 次";//A測試品
            Bwtext.Text = "連動測試完成 " + ACount + " 次";//B測試品
            Cwtext.Text = "連動測試完成 " + ACount + " 次";//C測試品
            Dwtext.Text = "連動測試完成 " + ACount + " 次";//D測試品
            Ewtext.Text = "連動測試完成 " + ACount + " 次";//E測試品
            workstr = "快速作動 完成";
            Adam6017a();
            Adam6017b();
            Adam6017c();
            Adam6017d();
            Adam6017e();
            Nowtimed(workstr);
            WorkStatus = workstr;
            Device = DeviceA1;
            Aps = Napa1.Text;
            WireCvsdata();
            Device = DeviceA2;
            Aps = Napa2.Text;
            WireCvsdata();
            Device = DeviceA3;
            Aps = Napa3.Text;
            WireCvsdata();
            Device = DeviceB1;
            Aps = Napb1.Text;
            WireCvsdata();
            Device = DeviceB2;
            Aps = Napb2.Text;
            WireCvsdata();
            Device = DeviceB3;
            Aps = Napb3.Text;
            WireCvsdata();
            Device = DeviceC1;
            Aps = Napc1.Text;
            WireCvsdata();
            Device = DeviceC2;
            Aps = Napc2.Text;
            WireCvsdata();
            Device = DeviceC3;
            Aps = Napc3.Text;
            WireCvsdata();
            Device = DeviceD1;
            Aps = Napd1.Text;
            WireCvsdata();
            Device = DeviceD2;
            Aps = Napd2.Text;
            WireCvsdata();
            Device = DeviceD3;
            Aps = Napd3.Text;
            WireCvsdata();
            Device = DeviceE1;
            Aps = Nape1.Text;
            WireCvsdata();
            Device = DeviceE2;
            Aps = Nape2.Text;
            WireCvsdata();
            Device = DeviceE3;
            Aps = Nape3.Text;
            WireCvsdata();
            SorDataTable();
            WorkNum = Convert.ToString(0);
            WorkUnm = Convert.ToString(0);
            ABox.Enabled = true;
            BBox.Enabled = true;
            CBox.Enabled = true;
            DBox.Enabled = true;
            EBox.Enabled = true;
            Bu_sp.Enabled = true;
            Bu_auto.Enabled = true;
            passAll = false;
            Bu_allclose.Enabled = false;
            Console.WriteLine("測試完成 passALL : " + passAll);    //檢查狀態
        }



        private void Rsa1_CheckedChanged(object sender, EventArgs e)
        {
            Bu_ack1.Visible = false;
            Bu_ack2.Visible = false;
            textBox2.Visible = false;
            textBox4.Visible = false;
            label3.Visible = false;
            label4.Visible = false;
        }

        private void Rsaa1_CheckedChanged(object sender, EventArgs e)
        {
            Bu_ack1.Visible = true;
            Bu_ack2.Visible = true;
            textBox2.Visible = true;
            textBox4.Visible = true;
            label3.Visible = true;
            label4.Visible = true;
        }

        private void Rsaa2_CheckedChanged(object sender, EventArgs e)
        {
            Bu_bck1.Visible = true;
            Bu_bck2.Visible = true;
            textBox29.Visible = true;
            textBox30.Visible = true;
            label28.Visible = true;
            label30.Visible = true;
        }

        private void Rsa2_CheckedChanged(object sender, EventArgs e)
        {
            Bu_bck1.Visible = false;
            Bu_bck2.Visible = false;
            textBox29.Visible = false;
            textBox30.Visible = false;
            label28.Visible = false;
            label30.Visible = false;
        }

        private void Rsaa3_CheckedChanged(object sender, EventArgs e)
        {
            Bu_cck1.Visible = true;
            Bu_cck2.Visible = true;
            textBox31.Visible = true;
            textBox33.Visible = true;
            label31.Visible = true;
            label33.Visible = true;
        }

        private void Rsa3_CheckedChanged(object sender, EventArgs e)
        {
            Bu_cck1.Visible = false;
            Bu_cck2.Visible = false;
            textBox31.Visible = false;
            textBox33.Visible = false;
            label31.Visible = false;
            label33.Visible = false;
        }

        private void Rsaa4_CheckedChanged(object sender, EventArgs e)
        {
            Bu_dck1.Visible = true;
            Bu_dck2.Visible = true;
            textBox22.Visible = true;
            textBox24.Visible = true;
            label22.Visible = true;
            label24.Visible = true;
        }

        private void Rsa4_CheckedChanged(object sender, EventArgs e)
        {
            Bu_dck1.Visible = false;
            Bu_dck2.Visible = false;
            textBox22.Visible = false;
            textBox24.Visible = false;
            label22.Visible = false;
            label24.Visible = false;
        }

        private void Rsaa5_CheckedChanged(object sender, EventArgs e)
        {

            Bu_eck1.Visible = true;
            Bu_eck2.Visible = true;
            textBox37.Visible = true;
            textBox35.Visible = true;
            label37.Visible = true;
            label35.Visible = true;
        }

        private void Rsa5_CheckedChanged(object sender, EventArgs e)
        {
            Bu_eck1.Visible = false;
            Bu_eck2.Visible = false;
            textBox37.Visible = false;
            textBox35.Visible = false;
            label37.Visible = false;
            label35.Visible = false;
        }


        public void TONA() //第四組執行內容的功能
        {
            Application.DoEvents();
            if (Rsa1.Checked == true)
            {
                if (passA == true) { return; }  //暫停跳出
                AWtimea = Convert.ToInt32(Wtime1.Text);   //快速作動秒數
                ACounta = Convert.ToInt32(textBox3.Text);   //動作次數;
                int rss = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rss == 0)
                {
                    Awtext.Text = "測試中";//A測試品           
                    Awtext.Refresh();
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 0, 1);
                    for (int ia = 1; ia <= ACounta; ia++)
                    {
                        if (passA == true) { return; }  //暫停跳出
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 1, 1);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 2, 0);
                        Thread.Sleep(AWtimea);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 1, 0);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 2, 1);
                        Thread.Sleep(AWtimea);
                        Awtext.Text = "測試中";//A測試品           
                        WorkNuma.Text = " 目前作動 " + ia + " 次";
                        Awtext.Refresh();
                        WorkNuma.Refresh();
                        Console.WriteLine("執行動作 1 passA = " + passA + " 目前作動 " + ia + " 次");//檢查狀態
                        WorkNum = Convert.ToString(ia);
                        TempTotalactiona = ia;
                    } //   for (int i = 1; i < ACount; i++)
                }//    if (rss == 0) 
                else
                {
                    if (passA == true) { return; }  //暫停跳出
                    zmcaux.ZAux_Close(ECI0064C);
                    zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                    Console.WriteLine("初始控制器 連接失敗");
                }
            }//   if (Rsa1.Checked == true)
            else
            {
                for (int ja = 1; ja <= AmunA; ja++) // 第一組保壓次數 D1保壓
                {
                    if (passA == true) { return; }  //暫停跳出
                    int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                    if (rssult == 0)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 0, 1);   //第一組 主電磁閥
                        zmcaux.ZAux_Close(ECI0064C);
                        Thread.Sleep(50);
                    }
                    else
                    {
                        zmcaux.ZAux_Close(ECI0064C);
                        zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                        Console.WriteLine("初始控制器 連接失敗");
                    }
                    for (int ia = 1; ia <= BmunA; ia++)//抓夾動作 次數 D1次數
                    {
                        Thread.Sleep(50);
                        if (passA == true) { return; }  //暫停跳出
                        if (Schang == "a")
                        {
                            if (passA == true) { return; }  //暫停跳出
                            int rss1 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                            if (rss1 == 0)
                            {
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 1, 1);  // 第一組                        
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 2, 0);  // 第一組                          
                                Thread.Sleep(WtimeA);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 1, 0); // 第一組
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 2, 1); // 第一組
                                Thread.Sleep(WtimeA);
                                Totalactiona = (ja - 1) * BmunA + ia;
                                Awtext.Text = "測試中";//A測試品           
                                WorkNuma.Text = " 目前作動 " + Totalactiona + " 次";
                                Awtext.Refresh();
                                WorkNuma.Refresh();
                                Console.WriteLine("A測試品 2 passA = " + passA + " 目前作動 " + Totalactiona + "次");//檢查狀態                               
                                zmcaux.ZAux_Close(ECI0064C);
                                Thread.Sleep(100);
                                TempTotalactiona = Totalactiona;
                                WorkNum = Convert.ToString(Totalactiona);

                            }
                            else
                            {
                                zmcaux.ZAux_Close(ECI0064C);
                                zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                                Console.WriteLine("A測試品 A控制器 " + Totalactiona + "次;控制器連接失敗");//檢查狀態  
                            }
                        }
                        else
                        {   // Schang="b"
                            if (passA == true) { return; }  //暫停跳出
                            int rss2 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                            if (rss2 == 0)
                            {
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 1, 0); // 第一組                           
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 2, 1); // 第一組                           
                                Thread.Sleep(WtimeA);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 1, 1); // 第一組                           
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 2, 0); // 第一組
                                Thread.Sleep(WtimeA);
                                Totalactiona = (ja - 1) * BmunA + ia;
                                Awtext.Text = "測試中";//A測試品           
                                WorkNuma.Text = " 目前作動 " + Totalactiona + " 次";
                                WorkNum = Convert.ToString(Totalactiona);
                                Awtext.Refresh();
                                WorkNuma.Refresh();
                                Console.WriteLine("A測試品 2  passA = " + passA + " 目前作動 " + Totalactiona + "次");//檢查狀態                                
                                zmcaux.ZAux_Close(ECI0064C);
                                Thread.Sleep(100);
                                TempTotalactiona = Totalactiona;
                                WorkNum = Convert.ToString(Totalactiona);
                            }
                            else
                            {
                                zmcaux.ZAux_Close(ECI0064C);
                                zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                                Console.WriteLine("A測試品  B控制器 " + Totalactiona + "次; 控制器連接失敗");//檢查狀態 
                            }
                        }
                    } //作動循環 for (int i = 1; i <= Bmun; i++)//抓夾動作 次數 D1次數
                    if (Schang == "a")  //左右停止 抓夾 A B
                        Schang = "b";
                    else Schang = "a";
                    //保壓動作 第一組主電磁閥

                    Awtext.Text = "測試保壓中 ; 第" + ja + " 次"; // A測試品
                    workstr = "測試品A 保壓";
                    WorkStatus = workstr;
                    Adam6017a();                  
                    Nowtimed(workstr);
                    WorkUnm = Convert.ToString(ja);
                    Device = DeviceA1;
                    WireCvsdata();
                    Device = DeviceA2;
                    WireCvsdata();
                    WorkStatus = workstr + "動作" + Schang.ToUpper();
                    Device = DeviceA3;
                    WireCvsdata();
                    Awtext.Refresh();
                    Console.WriteLine("A測試品  passA =" + passA + " 保壓中 第" + ja + " 次");//檢查狀態
                    Thread.Sleep(50);
                    if (passA == true) { return; }  //暫停跳出
                    int rss3 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);  //保壓等待時間
                    if (rss3 == 0)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 0, 0);
                        zmcaux.ZAux_Close(ECI0064C);
                        Thread.Sleep(TCountA);
                    }
                    else
                    {
                        zmcaux.ZAux_Close(ECI0064C);
                        zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                        Console.WriteLine("A測試品  保壓控制器 動作" + Totalactiona + "次; 控制器連接失敗");
                    }//  if (rss3 == 0)
                }    //保壓循理  for (int ja = 1; ja <= Amun; ja++) // 第一組保壓次數 D1保壓
            } //if (Rsa1.Checked == true) else 
            return;
        }//TONA


        public void TONB() //第四組執行內容的功能
        {
            Application.DoEvents();
            if (Rsa2.Checked == true)
            {
                AWtimeb = Convert.ToInt32(Wtime2.Text);   //快速作動秒數
                ACountb = Convert.ToInt32(textBox28.Text);   //動作次數;
                if (passB == true) { return; }  //暫停跳出
                int rss = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rss == 0)
                {
                    if (passB == true) { return; }  //暫停跳出
                    Bwtext.Text = "測試中";//B測試品           
                    Bwtext.Refresh();
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 4, 1);
                    for (int ib = 1; ib <= ACountb; ib++)
                    {
                        if (passB == true) { return; }  //暫停跳出
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 5, 1);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 6, 0);
                        Thread.Sleep(AWtimeb);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 5, 0);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 6, 1);
                        Thread.Sleep(AWtimeb);
                        Bwtext.Text = "測試中";//B測試品           
                        WorkNumb.Text = " 目前作動 " + ib + " 次";
                        WorkNum = Convert.ToString(ib);
                        TempTotalactionb = ib;
                        Bwtext.Refresh();
                        WorkNumb.Refresh();
                        Console.WriteLine("執行動作   passB = " + passB + " 目前作動 " + ib + " 次");//檢查狀態
                    } //   for (int i = 1; i < ACount; i++) 作動
                }
                else
                {
                    zmcaux.ZAux_Close(ECI0064C);
                    zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                    Console.WriteLine("B測試品 初始控制器 連接失敗");
                }
            }  //if (Rsa2.Checked == true)
            else
            {
                for (int jb = 1; jb <= AmunB; jb++) // 第二組保壓次數 D4保壓
                {
                    if (passB == true) { return; }  //暫停跳出
                    int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                    if (rssult == 0)
                    {
                        if (passB == true) { return; }  //暫停跳出
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 4, 1);
                        zmcaux.ZAux_Close(ECI0064C);
                        Thread.Sleep(100);
                    }
                    else
                    {
                        zmcaux.ZAux_Close(ECI0064C);
                        zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                        Console.WriteLine("B測試品 1 passB =" + passB + " 開始控制器 連接失敗");
                    }
                    for (int ib = 1; ib <= BmunB; ib++)//抓夾動作 次數 D5次數
                    {
                        if (passB == true) { return; }  //暫停跳出
                        if (Schang == "a")
                        {
                            if (passB == true) { return; }  //暫停跳出
                            int rss1 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                            if (rss1 == 0)
                            {
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 5, 1);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 6, 0);
                                Thread.Sleep(WtimeB);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 5, 0);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 6, 1);
                                Thread.Sleep(WtimeB);
                                Totalactionb = (jb - 1) * BmunB + ib;
                                Bwtext.Text = "測試中";//B測試品           
                                WorkNumb.Text = " 目前作動 " + Totalactionb + " 次";
                                WorkNum = Convert.ToString(Totalactionb);
                                TempTotalactionb = Totalactionb;
                                Bwtext.Refresh();
                                WorkNumb.Refresh();
                                Console.WriteLine("B測試品 2 passB = " + passB + " 目前作動 " + Totalactionb + "次");//檢查狀態                                
                                zmcaux.ZAux_Close(ECI0064C);
                                Thread.Sleep(100);
                            }
                            else
                            {
                                zmcaux.ZAux_Close(ECI0064C);
                                zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                                Console.WriteLine("B測試品  A控制器 作動" + Totalactionb + "次; 控制器連接失敗");//檢查狀態 
                            }
                        }
                        else
                        {
                            if (passB == true) { return; }  //暫停跳出
                            int rss2 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                            if (rss2 == 0)
                            {
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 5, 0);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 6, 1);
                                Thread.Sleep(WtimeB);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 5, 1);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 6, 0);
                                Thread.Sleep(WtimeB);
                                Totalactionb = (jb - 1) * BmunB + ib;
                                Bwtext.Text = "測試中";//B測試品           
                                WorkNumb.Text = " 目前作動 " + Totalactionb + " 次";
                                WorkNum = Convert.ToString(Totalactionb);
                                TempTotalactionb = Totalactionb;
                                Bwtext.Refresh();
                                WorkNumb.Refresh();
                                Console.WriteLine("B測試品 2 passB = " + passB + " 目前作動 " + Totalactionb + "次");//檢查狀態                        
                                zmcaux.ZAux_Close(ECI0064C);
                                Thread.Sleep(100);
                            }
                            else
                            {
                                zmcaux.ZAux_Close(ECI0064C);
                                zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                                Console.WriteLine("B測試品 B控制器 作動" + Totalactionb + "次; 控制器連接失敗");//檢查狀態 
                            }
                        }
                    }

                    if (Schang == "a")  //左右停止 抓夾 A B
                        Schang = "b";
                    else Schang = "a";
                    Bwtext.Text = "測試保壓中 ; 第" + jb + " 次"; //B測試品
                    workstr = "測試品B 保壓";                   
                    Adam6017b();                  
                    Nowtimed(workstr);
                    WorkStatus = workstr;
                    WorkUnm = Convert.ToString(jb);
                    Device = DeviceB1;
                    WireCvsdata();
                    Device = DeviceB2;
                    WireCvsdata();
                    WorkStatus = workstr + "動作" + Schang.ToUpper();
                    Device = DeviceB3;
                    WireCvsdata();
                    Bwtext.Refresh();
                    if (passB == true) { return; }  //暫停跳出
                                                    //保壓動作 第二組主電磁閥
                    Console.WriteLine("B測試品 2 passB = " + passB + "保壓中  第" + jb + " 次");//檢查狀態 
                    int rss3 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                    if (rss3 == 0)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 4, 0);
                        zmcaux.ZAux_Close(ECI0064C);
                        Thread.Sleep(TCountB);
                    }
                    else
                    {
                        zmcaux.ZAux_Close(ECI0064C);
                        zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                        Console.WriteLine("B測試品 保壓控制器 作動" + Totalactionb + "次;控制器連接失敗");//檢查狀態 
                    }
                }
            } //if (Rsa2.Checked == true) else 
            return;
        }//TONB

        public void TONC() //第四組執行內容的功能
        {
            Application.DoEvents();
            if (Rsa3.Checked == true)
            {
                AWtimec = Convert.ToInt32(Wtime3.Text);   //快速作動秒數
                ACountc = Convert.ToInt32(textBox32.Text);   //動作次數;
                if (passC == true) { return; }  //暫停跳出
                int rss = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rss == 0)
                {
                    Cwtext.Text = "測試開始"; //C測試品           
                    Cwtext.Refresh();
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 8, 1);
                    for (int ic = 1; ic < ACountc; ic++)
                    {
                        if (passC == true) { return; }  //暫停跳出
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 9, 1);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 10, 0);
                        Thread.Sleep(AWtimec);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 9, 0);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 10, 1);
                        Thread.Sleep(AWtimec);
                        Cwtext.Text = "測試中";//C測試品           
                        WorkNumc.Text = " 目前作動 " + ic + " 次";
                        WorkNum = Convert.ToString(ic);
                        TempTotalactionc = ic;
                        Cwtext.Refresh();
                        WorkNumc.Refresh();
                        Console.WriteLine("執行動作   passC = " + passC + " 目前作動 " + ic + " 次");//檢查狀態
                    } //   for (int i = 1; i < ACount; i++) 作動
                }
                else
                {
                    zmcaux.ZAux_Close(ECI0064C);
                    zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                    Console.WriteLine("C測試品 初始控制器 連接失敗");
                }
            }  //if (Rsa3.Checked == true)
            else
            {
                if (passC == true) { return; }  //暫停跳出
                for (int jc = 1; jc <= AmunC; jc++) // 第三組保壓次數 D8保壓
                {
                    if (passC == true) { return; }  //暫停跳出
                    int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                    if (rssult == 0)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 8, 1);
                        zmcaux.ZAux_Close(ECI0064C);
                        Thread.Sleep(100);
                    }
                    else
                    {
                        zmcaux.ZAux_Close(ECI0064C);
                        zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                        Console.WriteLine("C測試品 初始控制器 連接失敗");
                    }
                    for (int ic = 1; ic <= BmunC; ic++)//抓夾動作 次數 D1次數
                    {
                        if (passC == true) { return; }  //暫停跳出
                        if (Schang == "a")
                        {
                            int rss1 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                            if (rss1 == 0)
                            {
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 9, 1);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 10, 0);
                                Thread.Sleep(WtimeC);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 9, 0);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 10, 1);
                                Thread.Sleep(WtimeC);
                                Totalactionc = (jc - 1) * BmunC + ic;
                                Cwtext.Text = "測試中";//C測試品           
                                WorkNumc.Text = " 目前作動 " + Totalactionc + " 次";
                                WorkNum = Convert.ToString(Totalactionc);
                                TempTotalactionc = Totalactionc;
                                Cwtext.Refresh();
                                WorkNumc.Refresh();
                                zmcaux.ZAux_Close(ECI0064C);
                                Thread.Sleep(100);
                                Console.WriteLine("C測試品 2 passC = " + passC + " 目前作動 " + Totalactionc + " 次");
                                WorkNum = "" + Totalactionc;
                            }
                            else
                            {
                                zmcaux.ZAux_Close(ECI0064C);
                                zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                                Console.WriteLine("C測試品 A控制器 連接失敗第" + ic + "次");
                            }
                        }
                        else
                        {
                            if (passC == true) { return; }  //暫停跳出
                            int rss2 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                            if (rss2 == 0)
                            {
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 9, 0);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 10, 1);
                                Thread.Sleep(WtimeC);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 9, 1);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 10, 0);
                                Thread.Sleep(WtimeC);
                                Totalactionc = (jc - 1) * BmunC + ic;
                                Cwtext.Text = "測試中";//C測試品           
                                WorkNumc.Text = " 目前作動 " + Totalactionc + " 次";
                                WorkNum = Convert.ToString(Totalactionc);
                                TempTotalactionc = Totalactionc;
                                Cwtext.Refresh();
                                WorkNumc.Refresh();
                                zmcaux.ZAux_Close(ECI0064C);
                                Thread.Sleep(100);
                                Console.WriteLine("C測試品 2 passC = " + passC + " 目前作動 " + Totalactionc + " 次");
                            }
                            else
                            {
                                zmcaux.ZAux_Close(ECI0064C);
                                zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                                Console.WriteLine("C測試品 A控制器 連接失敗第" + ic + "次");
                            }
                        }
                    }
                    if (Schang == "a")  //左右停止 抓夾 A B
                        Schang = "b";
                    else Schang = "a";
                    Cwtext.Text = "測試C 保壓中 ; 第" + jc + " 次";//C測試品
                    workstr = "測試品C 保壓";                   
                    Adam6017c();                  
                    Nowtimed(workstr);
                    WorkUnm = Convert.ToString(jc);
                    WorkStatus = workstr;
                    Device = DeviceC1;
                    WireCvsdata();
                    Device = DeviceC2;
                    WireCvsdata();
                    WorkStatus = workstr + "動作" + Schang.ToUpper();
                    Device = DeviceC3;
                    WireCvsdata();
                    Cwtext.Refresh();
                    if (passC == true) { return; }  //暫停跳出                                                 
                    Console.WriteLine("C測試品 2 passC = " + passC + " 保壓中 ; 第" + jc + " 次");
                    //保壓動作 第三組主電磁閥
                    int rss3 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                    if (rss3 == 0)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 8, 0);
                        zmcaux.ZAux_Close(ECI0064C);
                        Thread.Sleep(TCountC);
                    }
                    else
                    {
                        zmcaux.ZAux_Close(ECI0064C);
                        zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                        Console.WriteLine("C測試品 保壓控制器 連接失敗第" + jc + "次");
                    }
                }
            } // if (Rsa3.Checked == true)
            return;
        }//TONC

        public void TOND() //第四組執行內容的功能
        {
            Application.DoEvents();
            if (Rsa4.Checked == true)
            {
                AWtimed = Convert.ToInt32(Wtime4.Text);   //快速作動秒數
                ACountd = Convert.ToInt32(textBox23.Text);   //動作次數;
                if (passD == true) { return; }  //暫停跳出
                int rss = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rss == 0)
                {
                    Dwtext.Text = "測試開始"; //D測試品           
                    Dwtext.Refresh();
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 12, 1);
                    for (int id = 1; id < ACountd; id++)
                    {
                        if (passD == true) { return; }  //暫停跳出
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 13, 1);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 14, 0);
                        Thread.Sleep(AWtimed);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 13, 0);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 14, 1);
                        Thread.Sleep(AWtimed);
                        Dwtext.Text = "測試中";//D測試品           
                        WorkNumd.Text = " 目前作動 " + id + " 次";
                        WorkNum = Convert.ToString(id);
                        TempTotalactiond = id;
                        Dwtext.Refresh();
                        WorkNumd.Refresh();
                        Console.WriteLine("執行動作   passD = " + passD + " 目前作動 " + id + " 次");//檢查狀態
                    } //   for (int id = 1; id < ACount; id++) 作動
                }
                else
                {
                    zmcaux.ZAux_Close(ECI0064C);
                    zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                    Console.WriteLine(" 初始控制器 連接失敗  ");
                }
            }  //if (Rsa4.Checked == true)
            else
            {
                for (int jd = 1; jd <= AmunD; jd++) // 第四組保壓次數 D8保壓
                {
                    if (passD == true) { return; }  //暫停跳出
                    int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                    if (rssult == 0)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 12, 1);
                        zmcaux.ZAux_Close(ECI0064C);
                        Thread.Sleep(100);
                    }
                    else
                    {
                        zmcaux.ZAux_Close(ECI0064C);
                        zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                        Console.WriteLine("D測試品  初始控制器 連接失敗");
                    }
                    for (int id = 1; id <= BmunD; id++)//抓夾動作 次數 D1次數
                    {
                        if (passD == true) { return; }  //暫停跳出
                        if (Schang == "a")
                        {
                            int rss1 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                            if (rss1 == 0)
                            {
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 13, 1);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 14, 0);
                                Thread.Sleep(WtimeD);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 13, 0);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 14, 1);
                                Thread.Sleep(WtimeD);
                                Totalactiond = (jd - 1) * BmunD + id;
                                Dwtext.Text = "測試中";//D測試品           
                                WorkNumd.Text = " 目前作動 " + Totalactiond + " 次";
                                WorkNum = Convert.ToString(Totalactiond);
                                TempTotalactiond = Totalactiond;
                                Dwtext.Refresh();
                                WorkNumd.Refresh();
                                Console.WriteLine("D測試品 2 passD = " + passD + " 目前作動 " + Totalactiond + " 次");
                                zmcaux.ZAux_Close(ECI0064C);
                                Thread.Sleep(100);
                            }
                            else
                            {
                                zmcaux.ZAux_Close(ECI0064C);
                                zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                                Console.WriteLine("D測試品  A組控制器 連接失敗");
                            }
                        }
                        else
                        {
                            if (passD == true) { return; }  //暫停跳出
                            int rss2 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                            if (rss2 == 0)
                            {
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 13, 0);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 14, 1);
                                Thread.Sleep(WtimeD);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 13, 1);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 14, 0);
                                Thread.Sleep(WtimeD);
                                Totalactiond = (jd - 1) * BmunD + id;
                                Dwtext.Text = "測試中";//D測試品           
                                WorkNumd.Text = " 目前作動 " + Totalactiond + " 次";
                                WorkNum = Convert.ToString(Totalactiond);
                                TempTotalactiond = Totalactiond;
                                Dwtext.Refresh();
                                WorkNumd.Refresh();
                                Console.WriteLine("D測試品 2 passD = " + passD + " 目前作動 " + Totalactiond + " 次");
                                zmcaux.ZAux_Close(ECI0064C);
                                Thread.Sleep(100);
                            }
                            else
                            {
                                zmcaux.ZAux_Close(ECI0064C);
                                zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                                Console.WriteLine("D測試品  B組控制器 連接失敗");
                            }
                        }
                    }
                    if (Schang == "a")  //左右停止 抓夾 A B
                        Schang = "b";
                    else Schang = "a";
                    Dwtext.Text = "測試品D 保壓中 ; 第" + jd + " 次";//D測試品
                    workstr = "測試品D 保壓";                   
                    Adam6017d();                   
                    Nowtimed(workstr);
                    WorkUnm = Convert.ToString(jd);
                    WorkStatus = workstr;
                    Device = DeviceD1;
                    WireCvsdata();
                    Device = DeviceD2;
                    WireCvsdata();
                    WorkStatus = workstr + "動作" + Schang.ToUpper();
                    Device = DeviceD3;
                    WireCvsdata();
                    Dwtext.Refresh();
                    if (passD == true) { return; }  //暫停跳出
                    //保壓動作 第四組主電磁閥
                    Console.WriteLine(" D測試品  2 passD = " + passD + " 保壓中  第" + jd + " 次");
                    int rss3 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                    if (rss3 == 0)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 12, 0);
                        zmcaux.ZAux_Close(ECI0064C);
                        Thread.Sleep(TCountD);
                    }
                    else
                    {
                        zmcaux.ZAux_Close(ECI0064C);
                        zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                        Console.WriteLine("D測試品  保壓控制器 連接失敗第" + jd + "次");
                    }
                }
            } // if (Rsa4.Checked == true) 
            return;
        } //TOND

        public void TONE() //第五組執行內容的功能
        {
            Application.DoEvents();
            if (Rsa5.Checked == true)
            {
                AWtimee = Convert.ToInt32(Wtime5.Text);   //快速作動秒數
                ACounte = Convert.ToInt32(textBox36.Text);   //動作次數;
                if (passE == true) { return; }//暫停跳出
                int rss = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rss == 0)
                {
                    Ewtext.Text = "測試開始"; //C測試品           
                    Ewtext.Refresh();
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 16, 1);
                    for (int ie = 1; ie < ACounte; ie++)
                    {
                        if (passE == true) { return; }//暫停跳出
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 17, 1);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 18, 0);
                        Thread.Sleep(AWtimee);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 17, 0);
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 18, 1);
                        Thread.Sleep(AWtimee);
                        Ewtext.Text = "測試中";//E測試品           
                        WorkNume.Text = " 目前作動 " + ie + " 次";
                        WorkNum = Convert.ToString(ie);
                        TempTotalactione = ie;
                        Ewtext.Refresh();
                        WorkNume.Refresh();
                        Console.WriteLine("E測試品 1 passE = " + passE + "  目前作動 " + ie + " 次");
                    } //   for (int i = 1; i < ACount; i++) 作動
                }
                else
                {
                    zmcaux.ZAux_Close(ECI0064C);
                    zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                    Console.WriteLine("初始控制器 連接失敗");
                }
            }  //if (Rsa3.Checked == true)
            else
            {
                if (passE == true) { return; }//暫停跳出
                for (int je = 1; je <= AmunE; je++) // 第五組保壓次數 D8保壓
                {
                    if (passE == true) { return; }//暫停跳出
                    int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                    if (rssult == 0)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 16, 1);
                        zmcaux.ZAux_Close(ECI0064C);
                        Thread.Sleep(50);
                    }
                    else
                    {
                        zmcaux.ZAux_Close(ECI0064C);
                        zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                        Console.WriteLine("E測試品  初始控制器 連接失敗");
                    }
                    for (int ie = 1; ie <= BmunE; ie++)//抓夾動作 次數 
                    {
                        if (passE == true) { return; }//暫停跳出
                        if (Schang == "a")
                        {
                            int rss1 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                            if (rss1 == 0)
                            {
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 17, 1);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 18, 0);
                                Thread.Sleep(WtimeE);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 17, 0);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 18, 1);
                                zmcaux.ZAux_Close(ECI0064C);
                                Thread.Sleep(WtimeE);
                                Totalactione = (je - 1) * BmunE + ie;
                                Ewtext.Text = "測試中";//E測試品           
                                WorkNume.Text = " 目前作動 " + Totalactione + " 次";
                                WorkNum = Convert.ToString(Totalactione);
                                TempTotalactione = Totalactione;
                                Ewtext.Refresh();
                                WorkNume.Refresh();
                                Console.WriteLine("E測試品 2 passE = " + passE + "   目前作動 " + Totalactione + " 次");
                            }
                            else
                            {
                                zmcaux.ZAux_Close(ECI0064C);
                                zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                                Console.WriteLine("E測試品  A組控制器 連接失敗");
                            }
                        }
                        else
                        {
                            if (passE == true) { return; }//暫停跳出
                            int rss1 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                            if (rss1 == 0)
                            {
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 17, 0);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 18, 1);
                                Thread.Sleep(WtimeE);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 17, 1);
                                zmcaux.ZAux_Direct_SetOp(ECI0064C, 18, 0);
                                Thread.Sleep(100);
                                zmcaux.ZAux_Close(ECI0064C);
                                Thread.Sleep(WtimeE);
                                Totalactione = (je - 1) * BmunE + ie;
                                Ewtext.Text = "測試中";//E測試品           
                                WorkNume.Text = " 目前作動 " + Totalactione + " 次";
                                WorkNum = Convert.ToString(Totalactione);
                                TempTotalactione = Totalactione;
                                Ewtext.Refresh();
                                WorkNume.Refresh();
                                Console.WriteLine("E測試品 2 passE = " + passE + "   目前作動 " + Totalactione + " 次");
                            }
                            else
                            {
                                zmcaux.ZAux_Close(ECI0064C);
                                zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                                Console.WriteLine("E測試品  B組控制器 連接失敗");
                            }
                        }
                    }

                    if (Schang == "a")  //左右停止 抓夾 A B
                        Schang = "b";
                    else Schang = "a";
                    Ewtext.Text = "測試保壓中 ; 第" + je + " 次";//E測試品
                    workstr = "測試品E 保壓";                  
                    Adam6017e();
                    Nowtimed(workstr);
                    WorkUnm = Convert.ToString(je);
                    Device = DeviceE1;
                    WireCvsdata();
                    Device = DeviceE2;
                    WireCvsdata();
                    WorkStatus = workstr + "動作" + Schang.ToUpper();
                    Device = DeviceE3;
                    WireCvsdata();
                    Ewtext.Refresh();
                    Console.WriteLine("E測試品 2 passE = " + passE + " 保壓中; 第" + je + " 次");
                    if (passE == true) { return; }//暫停跳出
                    //保壓動作 第五組主電磁閥
                    int rss3 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                    if (rss3 == 0)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, 16, 0);
                        zmcaux.ZAux_Close(ECI0064C);
                        Thread.Sleep(TCountE);
                    }
                    else
                    {
                        zmcaux.ZAux_Close(ECI0064C);
                        zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                        Console.WriteLine("E測試品  保壓控制器 連接失敗第" + Totalactione + "次");
                    }
                }
            }// if (Rsa5.Checked == true) 
            return;
        } //TONE


        public void TONAll() //第四組執行內容的功能
        {
            Application.DoEvents();
            Console.WriteLine("開始測試 passALL =" + passAll);//檢查狀態
            for (int j = 1; j <= Amun; j++) // 第一組保壓次數 D1保壓
            {
                if (passAll == true) { return; }  //暫停跳出
                Thread.Sleep(100);
                int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rssult == 0)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 0, 1);   //第一組 主電磁閥
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 4, 1);   //第二組 主電磁閥
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 8, 1);   //第三組 主電磁閥
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 12, 1);//第四組 主電磁閥
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 16, 1);//第五組 主電磁閥                 
                    zmcaux.ZAux_Close(ECI0064C);
                    Thread.Sleep(50);
                }
                else
                {
                    zmcaux.ZAux_Close(ECI0064C);
                    zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                    Console.WriteLine("測試執行控制卡連接失敗");//檢查狀態
                }
                for (int i = 1; i <= Bmun; i++)//抓夾動作 次數 D1次數
                {
                    if (passAll == true) { return; }  //暫停跳出
                    Thread.Sleep(100);
                    if (Schang == "a")
                    {
                        Thread.Sleep(50);
                        int rss1 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                        if (rss1 == 0)
                        {
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 1, 1);  // 第一組                        
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 2, 0);  // 第一組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 3, 1);  // 第一組
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 5, 1);  //第二組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 6, 0);  //第二組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 7, 1);  // 第二組
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 9, 1);   // 第三組 
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 10, 0); //第三組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 11, 1); //第三組
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 13, 1); //第四組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 14, 0);//第四組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 15, 1); //第四組
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 17, 0);//第五組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 18, 1);//第五組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 19, 0); //第五組
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 1, 0); // 第一組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 2, 1); // 第一組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 3, 0);  // 第一組
                            Totalaction = (j - 1) * Bmun + i;
                            Awtext.Text = "測試中";// A測試品           
                            WorkNuma.Text = " 目前作動 " + Totalaction + " 次";
                            Awtext.Refresh();
                            WorkNuma.Refresh();
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 5, 0); //第二組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 6, 1); //第二組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 7, 0); //第二組
                            Totalaction = (j - 1) * Bmun + i;
                            Bwtext.Text = "測試中";// B測試品           
                            WorkNumb.Text = " 目前作動 " + Totalaction + " 次";
                            Bwtext.Refresh();
                            WorkNumb.Refresh();
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 9, 0);   //第三組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 10, 1); //第三組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 11, 0);  //第三組
                            Totalaction = (j - 1) * Bmun + i;
                            Cwtext.Text = "測試中";// C測試品           
                            WorkNumc.Text = " 目前作動 " + Totalaction + " 次";
                            Cwtext.Refresh();
                            WorkNumc.Refresh();
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 13, 0);  //第四組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 14, 1);  //第四組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 15, 0);  //第四組
                            Totalaction = (j - 1) * Bmun + i;
                            Dwtext.Text = "測試中";// D測試品           
                            WorkNumd.Text = " 目前作動 " + Totalaction + " 次";
                            Dwtext.Refresh();
                            WorkNumd.Refresh();
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 17, 1);  //第五組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 18, 0);  //第五組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 19, 1);  //第五組
                            Totalaction = (j - 1) * Bmun + i;
                            Ewtext.Text = "測試中";// E測試品           
                            WorkNume.Text = " 目前作動 " + Totalaction + " 次";
                            WorkNum = Convert.ToString(Totalaction);
                            TempTotalaction = Totalaction;
                            Ewtext.Refresh();
                            WorkNume.Refresh();
                            Thread.Sleep(100);
                            zmcaux.ZAux_Close(ECI0064C);
                            Console.WriteLine("測試執行 2  passALL = " + passAll + " 目前作動 " + Totalaction + " 次");//檢查狀態
                            if (passAll == true) { return; }  //暫停跳出
                        }
                        else
                        {
                            zmcaux.ZAux_Close(ECI0064C);
                            zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                            Console.WriteLine("2  passALL = 測試執行控制卡連接失敗");//檢查狀態
                        }
                    }
                    else
                    {
                        if (passAll == true) { return; }  //暫停跳出
                        Thread.Sleep(100);
                        int rss2 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                        if (rss2 == 0)
                        {
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 1, 0); // 第一組                           
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 2, 1); // 第一組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 3, 0);  // 第一組                                         
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 5, 0);//第二組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 6, 1);//第二組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 7, 0); //第二組
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 9, 0);//第三組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 10, 1);//第三組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 11, 0); //第三組
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 13, 0); //第四組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 14, 1); //第四組 
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 15, 0); //第四組
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 17, 1);//第五組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 18, 0);//第五組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 19, 1); //第五組
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 1, 1); // 第一組                           
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 2, 0); // 第一組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 3, 1); // 第一組
                            Totalaction = (j - 1) * Bmun + i;
                            Awtext.Text = "測試中";// A測試品           
                            WorkNuma.Text = " 目前作動 " + Totalaction + " 次";
                            Awtext.Refresh();
                            WorkNuma.Refresh();
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 5, 1); //第二組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 6, 0); //第二組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 7, 1); //第二組
                            Totalaction = (j - 1) * Bmun + i;
                            Bwtext.Text = "測試中";// B測試品           
                            WorkNumb.Text = " 目前作動 " + Totalaction + " 次";
                            Bwtext.Refresh();
                            WorkNumb.Refresh();
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 9, 1);    //第三組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 10, 0); //第三組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 11, 1);  //第三組
                            Totalaction = (j - 1) * Bmun + i;
                            Cwtext.Text = "測試中";// C測試品           
                            WorkNumc.Text = " 目前作動 " + Totalaction + " 次";
                            Cwtext.Refresh();
                            WorkNumc.Refresh();
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 13, 1); //第四組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 14, 0); //第四組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 15, 1); //第四組
                            Totalaction = (j - 1) * Bmun + i;
                            Dwtext.Text = "測試中";// D測試品           
                            WorkNumd.Text = " 目前作動 " + Totalaction + " 次";
                            Dwtext.Refresh();
                            WorkNumd.Refresh();
                            Thread.Sleep(Wtime);
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 17, 0);  //第五組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 18, 1);  //第五組
                            zmcaux.ZAux_Direct_SetOp(ECI0064C, 19, 0);  //第五組
                            Totalaction = (j - 1) * Bmun + i;
                            Ewtext.Text = "測試中";// A測試品           
                            WorkNume.Text = " 目前作動 " + Totalaction + " 次";
                            WorkNum = Convert.ToString(Totalaction);
                            TempTotalaction = Totalaction;
                            Ewtext.Refresh();
                            WorkNume.Refresh();
                            Thread.Sleep(100);
                            zmcaux.ZAux_Close(ECI0064C);
                            Console.WriteLine("測試執行 2 passAll = " + passAll + " 目前作動 " + Totalaction + " 次");//檢查狀態
                            if (passAll == true) { return; }  //暫停跳出
                        }
                        else
                        {
                            zmcaux.ZAux_Close(ECI0064C);
                            zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                            Console.WriteLine("測試執行控制卡連接失敗");//檢查狀態
                        }
                    }
                }
                if (Schang == "a")  //左右停止 抓夾 A B
                    Schang = "b";
                else Schang = "a";
                //保壓動作 第一組主電磁閥
                Awtext.Text = "測試保壓中"; // A測試品
                Bwtext.Text = "測試保壓中"; //B測試品
                Cwtext.Text = "測試保壓中";//C測試品
                Dwtext.Text = "測試保壓中";//D測試品
                Ewtext.Text = "測試保壓中";//E測試品
                Awtext.Refresh();
                Bwtext.Refresh();
                Cwtext.Refresh();
                Dwtext.Refresh();
                Ewtext.Refresh();
                Adam6017a();
                Adam6017b();
                Adam6017c();
                Adam6017d();
                Adam6017e();
                WorkUnm = Convert.ToString(j);
                WorkNum = Convert.ToString(Totalaction);
                workstr = "自動保壓中";
                Nowtimed(workstr);
                WorkStatus = workstr;
                Device = DeviceA1;
                Aps = Napa1.Text;
                WireCvsdata();
                Device = DeviceA2;
                Aps = Napa2.Text;
                WireCvsdata();              
                Device = DeviceB1;
                Aps = Napb1.Text;
                WireCvsdata();
                Device = DeviceB2;
                Aps = Napb2.Text;
                WireCvsdata();               
                Device = DeviceC1;
                Aps = Napc1.Text;
                WireCvsdata();
                Device = DeviceC2;
                Aps = Napc2.Text;
                WireCvsdata();               
                Device = DeviceD1;
                Aps = Napd1.Text;
                WireCvsdata();
                Device = DeviceD2;
                Aps = Napd2.Text;
                WireCvsdata();             
                Device = DeviceE1;
                Aps = Nape1.Text;
                WireCvsdata();
                Device = DeviceE2;
                Aps = Nape2.Text;
                WireCvsdata();               
                WorkStatus = workstr + "動作" + Schang.ToUpper();
                Device = DeviceA3;
                Aps = Napa3.Text;
                WireCvsdata();
                Device = DeviceB3;
                Aps= Napb3.Text;
                WireCvsdata();
                Device = DeviceC3;
                Aps= Napc3.Text;
                WireCvsdata();
                Device = DeviceD3;
                Aps = Napd3.Text;
                WireCvsdata();
                Device = DeviceE3;
                Aps= Nape3.Text;
                WireCvsdata();
                SorDataTable();
                if (passAll == true) { return; }  //暫停跳出
                Thread.Sleep(50);
                Console.WriteLine("測試執行 2 passAll = " + passAll + "保壓動作 " + j + " 次");//檢查狀態
                int rss3 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);  //保壓等待時間
                if (rss3 == 0)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 0, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 4, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 8, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 12, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 16, 0);
                    zmcaux.ZAux_Close(ECI0064C);
                    Thread.Sleep(TCount);
                }
                else
                {
                    zmcaux.ZAux_Close(ECI0064C);
                    zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
                    Console.WriteLine("測試執行控制卡連接失敗");//檢查狀態
                }
            }
            return;
        }
        //快速執行內容的功能
        public void AutoStar()
        {
            Application.DoEvents();
            Console.WriteLine("1 passAll = " + passAll);//檢查狀態
            ACount = Convert.ToInt32(textBox3.Text);    //動作次數;
            AWtime = Convert.ToInt32(Wtime1.Text);    //自作動時間  AWtimea*100= Thread.Sleep(WtimeA)   Thread.Sleep(100)=10ms ;
            int rss = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
            if (rss == 0)
            {
                Cwtext.Text = "測試中";//C測試品           
                Cwtext.Refresh();
                zmcaux.ZAux_Direct_SetOp(ECI0064C, 0, 1);
                zmcaux.ZAux_Direct_SetOp(ECI0064C, 4, 1);
                zmcaux.ZAux_Direct_SetOp(ECI0064C, 8, 1);
                zmcaux.ZAux_Direct_SetOp(ECI0064C, 12, 1);
                zmcaux.ZAux_Direct_SetOp(ECI0064C, 16, 1);
                for (int i = 1; i <= ACount; i++)
                {
                    if (passAll == true) { return; }  //暫停跳出
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 1, 1);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 2, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 3, 1);
                    Thread.Sleep(AWtime);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 5, 1);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 6, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 7, 1);
                    Thread.Sleep(AWtime);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 9, 1);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 10, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 11, 1);
                    Thread.Sleep(AWtime);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 13, 1);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 14, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 15, 1);
                    Thread.Sleep(AWtime);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 17, 1);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 18, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 19, 1);
                    Thread.Sleep(AWtime);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 1, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 2, 1);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 3, 0);
                    Thread.Sleep(AWtime);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 5, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 6, 1);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 7, 0);
                    Thread.Sleep(AWtime);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 9, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 10, 1);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 11, 0);
                    Thread.Sleep(AWtime);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 13, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 14, 1);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 15, 0);
                    Thread.Sleep(AWtime);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 17, 0);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 18, 1);
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, 19, 0);
                    Thread.Sleep(AWtime);
                    Awtext.Text = "測試中";// A測試品           
                    WorkNuma.Text = " 目前作動 " + i + " 次";
                    Awtext.Refresh();
                    WorkNuma.Refresh();
                    Bwtext.Text = "測試中";// B測試品           
                    WorkNumb.Text = " 目前作動 " + i + " 次";
                    Bwtext.Refresh();
                    WorkNumb.Refresh();
                    Cwtext.Text = "測試中";// C測試品           
                    WorkNumc.Text = " 目前作動 " + i + " 次";
                    Cwtext.Refresh();
                    WorkNumc.Refresh(); ;
                    Dwtext.Text = "測試中";// D測試品           
                    WorkNumd.Text = " 目前作動 " + i + " 次";
                    Dwtext.Refresh();
                    WorkNumd.Refresh();
                    Ewtext.Text = "測試中";// A測試品           
                    WorkNume.Text = " 目前作動 " + i + " 次";
                    Ewtext.Refresh();
                    WorkNume.Refresh();
                    WorkNum = Convert.ToString(i);
                    WorkUnm = "";
                    Console.WriteLine("測試中 2  passALL = " + passAll + " 目前作動 " + i + " 次");    //檢查狀態
                    if (passAll == true) { return; }  //暫停跳出
                }
            }
            int rss1 = zmcaux.ZAux_OpenEth(ipaddr, out ECI0064C);
            if (rss1 == 0)
            {
                for (int i = 0; i < 31; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                }
            }
            return;
        }

        private void Bu_pas1_Click(object sender, EventArgs e)
        {
            passA = true;
            Bu_pas1.Enabled = false;
            workstr = "A暫停";
            Adam6017a();       
            Nowtimed(workstr);
            WorkNum = Convert.ToString(TempTotalactiona);
            WorkStatus = workstr;
            Device = DeviceA1;
            WireCvsdata();
            Device = DeviceA2;
            WireCvsdata();
            Device = DeviceA3;
            WireCvsdata();
            int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
            if (rssult == 0)
            {
                for (int i = 0; i <= 3; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                }
                zmcaux.ZAux_Close(ECI0064C);
                //  MessageBox.Show("全歸零成功");
            }
            return;
        }

        private void Bu_pas2_Click(object sender, EventArgs e)
        {
            passB = true;
            Bu_pas2.Enabled = false;
            workstr = "B暫停";
            Adam6017b();          
            Nowtimed(workstr);
            WorkNum = Convert.ToString(TempTotalactionb);
            WorkStatus = workstr;
            Device = DeviceB1;
            WireCvsdata();
            Device = DeviceB2;
            WireCvsdata();
            Device = DeviceB3;
            WireCvsdata();
            int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
            if (rssult == 0)
            {
                for (int i = 4; i <= 7; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                }
                zmcaux.ZAux_Close(ECI0064C);
                //  MessageBox.Show("全歸零成功");
            }
        }

        private void Bu_pas3_Click(object sender, EventArgs e)
        {
            passC = true;
            Bu_pas3.Enabled = false;
            workstr = "C暫停";        
            Adam6017c();          
            Nowtimed(workstr);
            WorkNum = Convert.ToString(TempTotalactionc);
            WorkStatus = workstr;
            Device = DeviceC1;
            WireCvsdata();
            Device = DeviceC2;
            WireCvsdata();
            Device = DeviceC3;
            WireCvsdata();
            int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
            if (rssult == 0)
            {
                for (int i = 8; i <= 11; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                }
                zmcaux.ZAux_Close(ECI0064C);
                //  MessageBox.Show("全歸零成功");
            }
            return;
        }

        private void Bu_pas4_Click(object sender, EventArgs e)
        {
            passD = true;
            Bu_pas4.Enabled = false;
            workstr = "D暫停";
            Adam6017d();
            Nowtimed(workstr);
            WorkNum = Convert.ToString(TempTotalactiond);
            WorkStatus = workstr;
            Device = DeviceD1;
            WireCvsdata();
            Device = DeviceD2;
            WireCvsdata();
            Device = DeviceD3;
            WireCvsdata();
            int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
            if (rssult == 0)
            {
                for (int i = 12; i <= 15; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                }
                zmcaux.ZAux_Close(ECI0064C);
                //  MessageBox.Show("全歸零成功");
            }
            return;
        }

        private void Bu_pas5_Click(object sender, EventArgs e)
        {
            passE = true;
            Bu_pas5.Enabled = false;
            workstr = "E暫停";         
            Adam6017e();
            Nowtimed(workstr);
            WorkNum = Convert.ToString(TempTotalactione);
            WorkStatus = workstr;
            Device = DeviceE1;
            WireCvsdata();
            Device = DeviceE2;
            WireCvsdata();
            Device = DeviceE3;
            WireCvsdata();
            int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
            if (rssult == 0)
            {
                for (int i = 16; i < 20; i++)
                {
                    zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                }
                zmcaux.ZAux_Close(ECI0064C);
                //  MessageBox.Show("全歸零成功");
            }
            return;
        }//Bu_pas5_Click  end                           

        private void LogIp_CheckedChanged(object sender, EventArgs e)
        {
            ipaddr = "127.0.0.1"; //本機測試
            return;
        }

        private void SerIp_CheckedChanged(object sender, EventArgs e)
        {
            ipaddr = "192.168.0.11"; // ECI0064 控制卡設定IP
            return;
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            if (LogIp.Checked)
            {
                ipaddr = "127.0.0.1"; //本機測試
            }
            else if (SerIp.Checked)
            {
                ipaddr = "192.168.0.11"; // ECI0064 控制卡設定IP
            }

             workstr = "程式啟動";
            Nowtimed(workstr);
            Device = "測試品";
            WorkStatus = workstr;
            WireCvsdata();
            // 讀入之前的資料
            TextA1.Text = Properties.Settings.Default.TextA1Data;
            TextA2.Text = Properties.Settings.Default.TextA2Data;
            TextA3.Text = Properties.Settings.Default.TextA3Data;
            TextB1.Text = Properties.Settings.Default.TextB1Data;
            TextB2.Text = Properties.Settings.Default.TextB2Data;
            TextB3.Text = Properties.Settings.Default.TextB3Data;
            TextC1.Text = Properties.Settings.Default.TextC1Data;
            TextC2.Text = Properties.Settings.Default.TextC2Data;
            TextC3.Text = Properties.Settings.Default.TextC3Data;
            TextD1.Text = Properties.Settings.Default.TextD1Data;
            TextD2.Text = Properties.Settings.Default.TextD2Data;
            TextD3.Text = Properties.Settings.Default.TextD3Data;
            TextE1.Text = Properties.Settings.Default.TextE1Data;
            TextE2.Text = Properties.Settings.Default.TextE2Data;
            TextE3.Text = Properties.Settings.Default.TextE3Data;

            //    this.FormClosing += Form1_FormClosing;
            this.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            return;
        }//Form1_load

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                //寫回原先設定值
                if (isDataChanged)
                {
                    Properties.Settings.Default.TextA1Data = TextA1.Text;
                    Properties.Settings.Default.TextA2Data = TextA2.Text;
                    Properties.Settings.Default.TextA3Data = TextA3.Text;
                    Properties.Settings.Default.TextB1Data = TextB1.Text;
                    Properties.Settings.Default.TextB2Data = TextB2.Text;
                    Properties.Settings.Default.TextB3Data = TextB3.Text;
                    Properties.Settings.Default.TextC1Data = TextC1.Text;
                    Properties.Settings.Default.TextC2Data = TextC2.Text;
                    Properties.Settings.Default.TextC3Data = TextC3.Text;
                    Properties.Settings.Default.TextD1Data = TextD1.Text;
                    Properties.Settings.Default.TextD2Data = TextD2.Text;
                    Properties.Settings.Default.TextD3Data = TextD3.Text;
                    Properties.Settings.Default.TextE1Data = TextE1.Text;
                    Properties.Settings.Default.TextE2Data = TextE2.Text;
                    Properties.Settings.Default.TextE3Data = TextE3.Text;
                    Properties.Settings.Default.Save();
                }
                //寫回原先設定值
                // 在這裡執行您想要的程式邏輯，例如保存資料或清理資源
                workstr = "程式結束";
                Nowtimed(workstr);
                WorkStatus = workstr;
                WireCvsdata();
                //全部歸零DO D0~D31 都輸出 為0
                int rssult = zmcaux.ZAux_OpenEth(ipaddr, out IntPtr ECI0064C);
                if (rssult == 0)
                {
                    for (int i = 0; i <= 31; i++)
                    {
                        zmcaux.ZAux_Direct_SetOp(ECI0064C, i, 0);
                    }
                    zmcaux.ZAux_Close(ECI0064C);
                    //  MessageBox.Show("全歸零成功");
                }
                else
                {
                    //       MessageBox.Show("歸零失敗");
                }
                // 資料回寫到CSV
                // 添加日誌記錄語句                
            }
            catch (Exception ex)
            {
                // 異常處理，例如記錄異常信息
                //寫回原先設定值
                if (isDataChanged)
                {
                    Properties.Settings.Default.TextA1Data = TextA1.Text;
                    Properties.Settings.Default.TextA2Data = TextA2.Text;
                    Properties.Settings.Default.TextA3Data = TextA3.Text;
                    Properties.Settings.Default.TextB1Data = TextB1.Text;
                    Properties.Settings.Default.TextB2Data = TextB2.Text;
                    Properties.Settings.Default.TextB3Data = TextB3.Text;
                    Properties.Settings.Default.TextC1Data = TextC1.Text;
                    Properties.Settings.Default.TextC2Data = TextC2.Text;
                    Properties.Settings.Default.TextC3Data = TextC3.Text;
                    Properties.Settings.Default.TextD1Data = TextD1.Text;
                    Properties.Settings.Default.TextD2Data = TextD2.Text;
                    Properties.Settings.Default.TextD3Data = TextD3.Text;
                    Properties.Settings.Default.TextE1Data = TextE1.Text;
                    Properties.Settings.Default.TextE2Data = TextE2.Text;
                    Properties.Settings.Default.TextE3Data = TextE3.Text;
                    Properties.Settings.Default.Save();
                }
                //寫回原先設定值
                workstr = "程式終止";
                Nowtimed(workstr);
                WorkStatus = workstr;
                WireCvsdata();
            }
            return;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void TextA1_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextA2_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextA3_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextB1_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextB2_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextB3_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextC1_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextC2_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextC3_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextD1_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextD2_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextD3_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextE1_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextE2_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void TextE3_TextChanged(object sender, EventArgs e)
        {
            isDataChanged = true;
        }

        private void Napa1_TextChanged(object sender, EventArgs e)
        {

        }
    }//Form1
}//ECI0064
