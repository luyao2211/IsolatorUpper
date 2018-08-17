using ImageSample_1;
using Newtonsoft.Json;
using OPCAutomation;
using SocketSerialTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace 隔离器自动运行服务测试
{
    class Program
    {
        //OPC参数
        static string strHostIP = "";
        static string strHostName = "";
        static OPCServer KepServer;
        static List<string> cmbServerName = new List<string>();
        static bool opc_connected = false;
        static List<string> listBox1 = new List<string>();
        static OPCGroups KepGroups;
        static OPCGroup KepGroup;
        static OPCItems KepItems;
        static OPCItems KepItemsEnvironment;
        static int itmHandleClient = 0;
        static int itmHandleServer = 0;
        static OPCItem KepItem;
        static string received = "";

        //串口参数
        static List<string> sl = new List<string>();
        static private System.IO.Ports.SerialPort serialPort = new System.IO.Ports.SerialPort();

        //监听变量
        static int? selectedAddress = null;
        //正在执行的任务
        static List<TestResultInfo> Processing = new List<TestResultInfo>();
        //正在执行的操作
        static List<OperationInfo> OperationNow = new List<OperationInfo>();
        //下一步操作
        static List<string> OperationNext = new List<string>();
        //二维码txt信息存储地址
        static string ErweiDir = "D:\\数据追溯\\";
        //拆垛端托盘数
        static int Tuopan = 0;

        static void Main(string[] args)
        {
            GetLocalServer();
            btnConnLocalServer_Click();
            init_Serial_List();
            serialPort.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(serialPort_DataReceived);

            while (0 != 1)
            {
                if (Processing.Count == 0)
                {
                    string strURL = "http://121.43.107.106:8063/Api/v1/Result/ResTestResultGetResultInfosByAnyProperty";
                    string strBody = "{\"TestId\": null,\"ObjectNo\": null,\"ObjCompany\": null,\"ObjIncuSeq\": null,\"TestType\": null,\"TestStand\": null,\"TestEquip\": null,\"TestEquip2\": null,\"Description\": null,\"ProcessStartS\": null,\"ProcessStartE\": null,\"ProcessEndS\": \"0001-01-01T00:00:00\",\"ProcessEndE\": null,\"CollectStartS\": null,\"CollectStartE\": null,\"CollectEndS\": null,\"CollectEndE\": null,\"TestTimeS\": null,\"TestTimeE\": null,\"TestResult\": null,\"TestPeople\": null,\"TestPeople2\": null,\"ReStatus\": null,\"RePeople\": null,\"ReTimeS\": null,\"ReTimeE\": null,\"ReDateTimeS\": null,\"ReDateTimeE\": null,\"ReTerminalIP\": null,\"ReTerminalName\": null,\"ReUserId\": null,\"ReIdentify\": null,\"FormerStep\": null,\"NowStep\": null,\"LaterStep\": null,\"GetObjectNo\": 1,\"GetObjCompany\": 1,\"GetObjIncuSeq\": 1,\"GetTestType\": 1,\"GetTestStand\": 1,\"GetTestEquip\": 1,\"GetTestEquip2\": 1,\"GetDescription\": 1,\"GetProcessStart\": 1,\"GetProcessEnd\": 1,\"GetCollectStart\": 1,\"GetCollectEnd\": 1,\"GetTestTime\": 1,\"GetTestResult\": 1,\"GetTestPeople\": 1,\"GetTestPeople2\": 1,\"GetReStatus\": 1,\"GetRePeople\": 1,\"GetReTime\": 1,\"GetRevisionInfo\": 1,\"GetFormerStep\": 1,\"GetNowStep\": 1,\"GetLaterStep\": 1}";
                    List<string> strTests = HttpPost(strURL, strBody).Split('}').ToList();
                    strTests.Remove("]");
                    if (strTests[0] != "[]")
                    {
                        for (int i = 0; i < strTests.Count; i++)
                        {
                            TestResultInfo testResult = JsonHelper.DeserializeJsonToObject<TestResultInfo>(strTests[i].Substring(1, strTests[i].Length - 1) + "}");
                            Processing.Add(testResult);
                            strURL = "http://121.43.107.106:8063/Api/v1/Operation/OpEquipmentSetData";
                            OperationNow.Add(new OperationInfo
                            {
                                EquipmentId = "Iso_Process",
                                OperationTime = DateTime.Now,
                                OperationCode = "OP001",
                                OperationValue = "1",
                                OperationResult = "",
                                TerminalIP = "127.0.0.1",
                                TerminalName = Dns.GetHostName(),
                                revUserId = "SHANGWEIJI"
                            });
                            OperationNext.Add("OP002");
                            Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(OperationNow[i])));
                        }
                    }
                    Tuopan = 3;
                }
                else
                {
                    if (OperationNow[0].OperationCode == "OP001")
                    {
                        open_Serial(0);
                        while (1 != 0)
                        {
                            Thread.Sleep(1000);
                            if (received != "")
                            {
                                //上传二维码信息
                                FtpHelper ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                string filePath = ErweiDir + received.Split('\r').ToList()[0] + ".txt";
                                if (!System.IO.File.Exists(filePath))
                                {
                                    FileStream fs1 = new FileStream(filePath, FileMode.Create, FileAccess.Write);//创建写入文件 
                                    StreamWriter sw = new StreamWriter(fs1);
                                    sw.WriteLine(Processing[0].TestId + "_1");//开始写入值
                                    sw.Close();
                                    fs1.Close();
                                }
                                FileInfo erwei = new FileInfo(filePath);
                                bool shangchuan = ftpHelper.Upload(erwei, "\\SterilityWebAPI\\erwei\\" + received.Split('\r').ToList()[0] + ".txt");
                                received = "";
                                close_Serial();
                                GetNextOperation(0);
                                break;
                            }
                        }
                    }
                    if (OperationNow[0].OperationCode == "OP002")
                    {
                        open_Serial(0);
                        while (1 != 0)
                        {
                            Thread.Sleep(1000);
                            if (received != "")
                            {
                                //上传二维码信息
                                FtpHelper ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                string filePath = ErweiDir + received.Split('\r').ToList()[0] + ".txt";
                                if (!System.IO.File.Exists(filePath))
                                {
                                    FileStream fs1 = new FileStream(filePath, FileMode.Create, FileAccess.Write);//创建写入文件 
                                    StreamWriter sw = new StreamWriter(fs1);
                                    sw.WriteLine(Processing[0].TestId + "_2");//开始写入值
                                    sw.Close();
                                    fs1.Close();
                                }
                                FileInfo erwei = new FileInfo(filePath);
                                bool shangchuan = ftpHelper.Upload(erwei, "\\SterilityWebAPI\\erwei\\" + received.Split('\r').ToList()[0] + ".txt");
                                received = "";
                                close_Serial();
                                GetNextOperation(0);
                                break;
                            }
                        }
                    }
                    if (OperationNow[0].OperationCode == "OP003")
                    {
                        open_Serial(0);
                        while (1 != 0)
                        {
                            Thread.Sleep(1000);
                            if (received != "")
                            {
                                //上传二维码信息
                                FtpHelper ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                string filePath = ErweiDir + received.Split('\r').ToList()[0] + ".txt";
                                if (!System.IO.File.Exists(filePath))
                                {
                                    FileStream fs1 = new FileStream(filePath, FileMode.Create, FileAccess.Write);//创建写入文件 
                                    StreamWriter sw = new StreamWriter(fs1);
                                    sw.WriteLine(Processing[0].TestId + "_3");//开始写入值
                                    sw.Close();
                                    fs1.Close();
                                }
                                FileInfo erwei = new FileInfo(filePath);
                                bool shangchuan = ftpHelper.Upload(erwei, "\\SterilityWebAPI\\erwei\\" + received.Split('\r').ToList()[0] + ".txt");
                                received = "";
                                close_Serial();
                                GetNextOperation(0);
                                if (Processing.Count == 1)
                                {
                                    listBox1_SelectedIndexChanged("通道 1.设备 1.入料段线体有料", false, 1234);
                                }
                                else
                                {
                                    if (OperationNow[1].OperationCode == "OP001")
                                    {
                                        open_Serial(0);
                                        while (1 != 0)
                                        {
                                            Thread.Sleep(1000);
                                            if (received != "")
                                            {
                                                //上传二维码信息
                                                ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                                filePath = ErweiDir + received.Split('\r').ToList()[0] + ".txt";
                                                if (!System.IO.File.Exists(filePath))
                                                {
                                                    FileStream fs1 = new FileStream(filePath, FileMode.Create, FileAccess.Write);//创建写入文件 
                                                    StreamWriter sw = new StreamWriter(fs1);
                                                    sw.WriteLine(Processing[1].TestId + "_1");//开始写入值
                                                    sw.Close();
                                                    fs1.Close();
                                                }
                                                erwei = new FileInfo(filePath);
                                                shangchuan = ftpHelper.Upload(erwei, "\\SterilityWebAPI\\erwei\\" + received.Split('\r').ToList()[0] + ".txt");
                                                received = "";
                                                close_Serial();
                                                GetNextOperation(1);
                                                break;
                                            }
                                        }
                                    }
                                    if (OperationNow[1].OperationCode == "OP002")
                                    {
                                        open_Serial(0);
                                        while (1 != 0)
                                        {
                                            Thread.Sleep(1000);
                                            if (received != "")
                                            {
                                                //上传二维码信息
                                                ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                                filePath = ErweiDir + received.Split('\r').ToList()[0] + ".txt";
                                                if (!System.IO.File.Exists(filePath))
                                                {
                                                    FileStream fs1 = new FileStream(filePath, FileMode.Create, FileAccess.Write);//创建写入文件 
                                                    StreamWriter sw = new StreamWriter(fs1);
                                                    sw.WriteLine(Processing[1].TestId + "_2");//开始写入值
                                                    sw.Close();
                                                    fs1.Close();
                                                }
                                                erwei = new FileInfo(filePath);
                                                shangchuan = ftpHelper.Upload(erwei, "\\SterilityWebAPI\\erwei\\" + received.Split('\r').ToList()[0] + ".txt");
                                                received = "";
                                                close_Serial();
                                                GetNextOperation(1);
                                                break;
                                            }
                                        }
                                    }
                                    if (OperationNow[1].OperationCode == "OP003")
                                    {
                                        open_Serial(0);
                                        while (1 != 0)
                                        {
                                            Thread.Sleep(1000);
                                            if (received != "")
                                            {
                                                //上传二维码信息
                                                ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                                filePath = ErweiDir + received.Split('\r').ToList()[0] + ".txt";
                                                if (!System.IO.File.Exists(filePath))
                                                {
                                                    FileStream fs1 = new FileStream(filePath, FileMode.Create, FileAccess.Write);//创建写入文件 
                                                    StreamWriter sw = new StreamWriter(fs1);
                                                    sw.WriteLine(Processing[1].TestId + "_3");//开始写入值
                                                    sw.Close();
                                                    fs1.Close();
                                                }
                                                erwei = new FileInfo(filePath);
                                                shangchuan = ftpHelper.Upload(erwei, "\\SterilityWebAPI\\erwei\\" + received.Split('\r').ToList()[0] + ".txt");
                                                received = "";
                                                close_Serial();
                                                GetNextOperation(1);
                                                break;
                                            }
                                        }
                                    }
                                    if (OperationNow[2].OperationCode == "OP001")
                                    {
                                        open_Serial(0);
                                        while (1 != 0)
                                        {
                                            Thread.Sleep(1000);
                                            if (received != "")
                                            {
                                                //上传二维码信息
                                                ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                                filePath = ErweiDir + received.Split('\r').ToList()[0] + ".txt";
                                                if (!System.IO.File.Exists(filePath))
                                                {
                                                    FileStream fs1 = new FileStream(filePath, FileMode.Create, FileAccess.Write);//创建写入文件 
                                                    StreamWriter sw = new StreamWriter(fs1);
                                                    sw.WriteLine(Processing[2].TestId + "_1");//开始写入值
                                                    sw.Close();
                                                    fs1.Close();
                                                }
                                                erwei = new FileInfo(filePath);
                                                shangchuan = ftpHelper.Upload(erwei, "\\SterilityWebAPI\\erwei\\" + received.Split('\r').ToList()[0] + ".txt");
                                                received = "";
                                                close_Serial();
                                                GetNextOperation(2);
                                                break;
                                            }
                                        }
                                    }
                                    if (OperationNow[2].OperationCode == "OP002")
                                    {
                                        open_Serial(0);
                                        while (1 != 0)
                                        {
                                            Thread.Sleep(1000);
                                            if (received != "")
                                            {
                                                //上传二维码信息
                                                ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                                filePath = ErweiDir + received.Split('\r').ToList()[0] + ".txt";
                                                if (!System.IO.File.Exists(filePath))
                                                {
                                                    FileStream fs1 = new FileStream(filePath, FileMode.Create, FileAccess.Write);//创建写入文件 
                                                    StreamWriter sw = new StreamWriter(fs1);
                                                    sw.WriteLine(Processing[2].TestId + "_2");//开始写入值
                                                    sw.Close();
                                                    fs1.Close();
                                                }
                                                erwei = new FileInfo(filePath);
                                                shangchuan = ftpHelper.Upload(erwei, "\\SterilityWebAPI\\erwei\\" + received.Split('\r').ToList()[0] + ".txt");
                                                received = "";
                                                close_Serial();
                                                GetNextOperation(2);
                                                break;
                                            }
                                        }
                                    }
                                    if (OperationNow[2].OperationCode == "OP003")
                                    {
                                        open_Serial(0);
                                        while (1 != 0)
                                        {
                                            Thread.Sleep(1000);
                                            if (received != "")
                                            {
                                                //上传二维码信息
                                                ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                                filePath = ErweiDir + received.Split('\r').ToList()[0] + ".txt";
                                                if (!System.IO.File.Exists(filePath))
                                                {
                                                    FileStream fs1 = new FileStream(filePath, FileMode.Create, FileAccess.Write);//创建写入文件 
                                                    StreamWriter sw = new StreamWriter(fs1);
                                                    sw.WriteLine(Processing[2].TestId + "_3");//开始写入值
                                                    sw.Close();
                                                    fs1.Close();
                                                }
                                                erwei = new FileInfo(filePath);
                                                shangchuan = ftpHelper.Upload(erwei, "\\SterilityWebAPI\\erwei\\" + received.Split('\r').ToList()[0] + ".txt");
                                                received = "";
                                                close_Serial();
                                                GetNextOperation(2);
                                                listBox1_SelectedIndexChanged("通道 1.设备 1.入料段线体有料", false, 1234);
                                                break;
                                            }
                                        }
                                    }
                                }
                                break;
                            }
                        }
                    }
                }
                Thread.Sleep(1000);
            }
        }

        /// <summary>
        /// 获取本地OPC服务器列表
        /// </summary>
        static private void GetLocalServer()
        {
            //获取本地计算机IP,计算机名称
            IPHostEntry IPHost = Dns.Resolve(Environment.MachineName);
            if (IPHost.AddressList.Length > 0)
            {
                strHostIP = IPHost.AddressList[0].ToString();
            }
            else
            {
                return;
            }
            //通过IP来获取计算机名称，可用在局域网内
            IPHostEntry ipHostEntry = Dns.GetHostByAddress(strHostIP);
            strHostName = ipHostEntry.HostName.ToString();
            //获取本地计算机上的OPCServerName
            try
            {
                KepServer = new OPCServer();
                object serverList = KepServer.GetOPCServers(strHostName);

                foreach (string turn in (Array)serverList)
                {
                    cmbServerName.Add(turn);
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("枚举本地OPC服务器出错：" + err.Message);
            }
            return;
        }

        /// <summary>
        /// 连接ＯＰＣ服务器
        /// </summary>
        static private void btnConnLocalServer_Click()
        {
            try
            {
                if (!ConnectRemoteServer("", cmbServerName[1]))
                {
                    return;
                }
                opc_connected = true;
                GetServerInfo();
                RecurBrowse(KepServer.CreateBrowser());
                if (!CreateGroup())
                {
                    return;
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("初始化出错：" + err.Message);
            }
            return;
        }

        /// <summary>
        /// 连接OPC服务器
        /// </summary>
        /// <param name="remoteServerIP">OPCServerIP</param>
        /// <param name="remoteServerName">OPCServer名称</param>
        static private bool ConnectRemoteServer(string remoteServerIP, string remoteServerName)
        {
            try
            {
                KepServer.Connect(remoteServerName, remoteServerIP);
                if (KepServer.ServerState == (int)OPCServerState.OPCRunning)
                {
                    Console.Write("已连接到-" + KepServer.ServerName + "\t");
                }
                else
                {
                    //这里你可以根据返回的状态来自定义显示信息，请查看自动化接口API文档
                    Console.Write("状态：" + KepServer.ServerState.ToString() + "\t");
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("连接远程服务器出现错误：" + err.Message);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 获取服务器信息，并显示在窗体状态栏上
        /// </summary>
        static private void GetServerInfo()
        {
            Console.Write("开始时间:" + KepServer.StartTime.ToString() + "\t");
            Console.WriteLine("版本:" + KepServer.MajorVersion.ToString() + "." + KepServer.MinorVersion.ToString() + "." + KepServer.BuildNumber.ToString());
            return;
        }

        /// <summary>
        /// 列出OPC服务器中所有节点
        /// </summary>
        /// <param name="oPCBrowser"></param>
        static private void RecurBrowse(OPCBrowser oPCBrowser)
        {
            //展开分支
            oPCBrowser.ShowBranches();
            //展开叶子
            oPCBrowser.ShowLeafs(true);
            foreach (object turn in oPCBrowser)
            {
                listBox1.Add(turn.ToString());
            }
            return;
        }

        /// <summary>
        /// 创建组
        /// </summary>
        static private bool CreateGroup()
        {
            try
            {
                KepGroups = KepServer.OPCGroups;
                KepGroup = KepGroups.Add("OPCDOTNETGROUP");
                SetGroupProperty();
                KepGroup.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(KepGroup_DataChange);
                KepGroup.AsyncWriteComplete += new DIOPCGroupEvent_AsyncWriteCompleteEventHandler(KepGroup_AsyncWriteComplete);
                KepItems = KepGroup.OPCItems;
                KepItemsEnvironment = KepGroup.OPCItems;
            }
            catch (Exception err)
            {
                Console.WriteLine("创建组出现错误：" + err.Message);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 设置组属性
        /// </summary>
        static private void SetGroupProperty()
        {
            KepServer.OPCGroups.DefaultGroupIsActive = Convert.ToBoolean("true");
            KepServer.OPCGroups.DefaultGroupDeadband = Convert.ToInt32("0");
            KepGroup.UpdateRate = Convert.ToInt32("250");
            KepGroup.IsActive = Convert.ToBoolean("true");
            KepGroup.IsSubscribed = Convert.ToBoolean("true");
            return;
        }

        /// <summary>
        /// 每当项数据有变化时执行的事件
        /// </summary>
        /// <param name="TransactionID">处理ID</param>
        /// <param name="NumItems">项个数</param>
        /// <param name="ClientHandles">项客户端句柄</param>
        /// <param name="ItemValues">TAG值</param>
        /// <param name="Qualities">品质</param>
        /// <param name="TimeStamps">时间戳</param>
        static void KepGroup_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            if (ItemValues.GetValue(1).ToString() == "True")
            {
                switch (selectedAddress)
                {
                    case 120:
                        for (int i = 0; i < OperationNow.Count; i++)
                        {
                            GetNextOperation(i);
                        }
                        listBox1_SelectedIndexChanged("通道 1.设备 2.进料门未关", false, 1234);
                        break;
                    case 29:
                        for (int i = 0; i < OperationNow.Count; i++)
                        {
                            GetNextOperation(i);
                        }
                        listBox1_SelectedIndexChanged("通道 1.设备 2.左传递门未关", false, 1234);
                        break;
                    case 0:
                        for (int i = 0; i < OperationNow.Count; i++)
                        {
                            GetNextOperation(i);
                        }
                        listBox1_SelectedIndexChanged("通道 1.设备 1.拆垛段线体上有料", false, 1234);
                        break;
                    case 152:
                        for (int i = 0; i < OperationNow.Count; i++)
                        {
                            GetNextOperation(i);
                        }
                        listBox1_SelectedIndexChanged("通道 1.设备 2.左传递门未关", false, 1234);
                        break;
                    case 115:
                        if (Processing.Count == 3)
                        {
                            GetNextOperation(3 - Tuopan);
                        }
                        else
                        {
                            GetNextOperation(0);
                        }
                        Tuopan = Tuopan - 1;
                        break;
                    case 128:
                        for (int i = 0; i < OperationNow.Count; i++)
                        {
                            GetNextOperation(i);
                        }
                        listBox1_SelectedIndexChanged("通道 1.设备 2.右传递门未关", false, 1234);
                        break;
                    case 2:
                        for (int i = 0; i < OperationNow.Count; i++)
                        {
                            GetNextOperation(i);
                        }
                        listBox1_SelectedIndexChanged("通道 1.设备 1.出料段线体有料", false, 1234);
                        break;
                    case 146:
                        for (int i = 0; i < OperationNow.Count; i++)
                        {
                            GetNextOperation(i);
                        }
                        listBox1_SelectedIndexChanged("通道 1.设备 2.右传递门未关", false, 1234);
                        break;
                }
            }
            if (ItemValues.GetValue(1).ToString() == "false")
            {
                switch (selectedAddress)
                {
                    case 27:
                        for (int i = 0; i < OperationNow.Count; i++)
                        {
                            GetNextOperation(i);
                        }
                        listBox1_SelectedIndexChanged("通道 1.设备 2.进料舱自动运行", false, 1234);
                        break;
                    case 0:
                        for (int i = 0; i < OperationNow.Count; i++)
                        {
                            GetNextOperation(i);
                        }
                        listBox1_SelectedIndexChanged("通道 1.设备 1.生产段有料", false, 1234);
                        break;
                    case 115:
                        if (Processing.Count == 3)
                        {
                            GetNextOperation(2 - Tuopan);
                        }
                        else
                        {
                            GetNextOperation(0);
                        }
                        if (Tuopan == 0)
                        {
                            listBox1_SelectedIndexChanged("通道 1.设备 1.堆垛段线体上有料", false, 1234);
                        }
                        break;
                    case 2:
                        for (int i = 0; i < OperationNow.Count; i++)
                        {
                            GetNextOperation(i);
                        }
                        listBox1_SelectedIndexChanged("通道 1.设备 1.出料段线体有料", false, 1234);
                        break;
                    case 146:
                        for (int i = 0; i < OperationNow.Count; i++)
                        {
                            EndOperation(i);
                        }
                        Processing = new List<TestResultInfo>();
                        OperationNow = new List<OperationInfo>();
                        OperationNext = new List<string>();
                        listBox1_SelectedIndexChanged("通道 1.设备 2.出料门未关", false, 1234);
                        break;
                    case 48:
                        //出料舱灭菌
                        break;
                }
            }
        }

        /// <summary>
        /// 写入TAG值时执行的事件
        /// </summary>
        /// <param name="TransactionID"></param>
        /// <param name="NumItems"></param>
        /// <param name="ClientHandles"></param>
        /// <param name="Errors"></param>
        static void KepGroup_AsyncWriteComplete(int TransactionID, int NumItems, ref Array ClientHandles, ref Array Errors)
        {
            for (int i = 1; i <= NumItems; i++)
            {
                Console.WriteLine("Tran:" + TransactionID.ToString() + "   CH:" + ClientHandles.GetValue(i).ToString() + "   Error:" + Errors.GetValue(i).ToString());
            }
            return;
        }

        /// <summary>
        /// 选择列表项时处理的事情
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static private void listBox1_SelectedIndexChanged(string tag, bool ifEnv, int clientNum)
        {
            try
            {
                if (ifEnv == true)
                {
                    itmHandleClient = clientNum;
                    KepItem = KepItemsEnvironment.AddItem(tag, itmHandleClient);
                    itmHandleServer = KepItem.ServerHandle;
                }
                if (ifEnv == false)
                {
                    if (itmHandleClient != 0)
                    {
                        Array Errors;
                        OPCItem bItem = KepItems.GetOPCItem(itmHandleServer);
                        //注：OPC中以1为数组的基数
                        int[] temp = new int[2] { 0, bItem.ServerHandle };
                        Array serverHandle = (Array)temp;
                        //移除上一次选择的项
                        KepItems.Remove(KepItems.Count, ref serverHandle, out Errors);
                    }
                    itmHandleClient = 1234;
                    KepItem = KepItems.AddItem(tag, itmHandleClient);
                    itmHandleServer = KepItem.ServerHandle;
                    selectedAddress = listBox1.FindIndex(s => s == tag);
                }
            }
            catch (Exception err)
            {
                //没有任何权限的项，都是OPC服务器保留的系统项，此处可不做处理。
                itmHandleClient = 0;
                Console.WriteLine("此项为系统保留项:" + err.Message, "提示信息");
            }
            return;
        }

        /// <summary>
        /// 【按钮】写入
        /// </summary>
        static private void btnWrite_Click(string outcome)
        {
            OPCItem bItem = KepItems.GetOPCItem(itmHandleServer);
            int[] temp = new int[2] { 0, bItem.ServerHandle };
            Array serverHandles = (Array)temp;
            object[] valueTemp = new object[2] { "", outcome };
            Array values = (Array)valueTemp;
            Array Errors;
            int cancelID;
            KepGroup.AsyncWrite(1, ref serverHandles, ref values, out Errors, 2009, out cancelID);
            //KepItem.Write(txtWriteTagValue.Text);//这句也可以写入，但并不触发写入事件
            GC.Collect();
            return;
        }

        /// <summary>
        /// 获取串口列表
        /// </summary>
        static private void init_Serial_List()
        {
            sl = SerialPortTool.GetSerialPortList();
            if (sl == null)
            {
                Console.WriteLine("读取串口列表失败");
                return;
            }
            return;
        }

        /// <summary>
        /// 读取串口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static private void serialPort_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            byte[] buffer = new byte[1024];
            int n = serialPort.Read(buffer, 0, 1024);
            received = System.Text.Encoding.UTF8.GetString(buffer, 0, n);
            return;
        }

        /// <summary>
        /// 关闭串口
        /// </summary>
        static private void close_Serial()
        {
            if (serialPort.IsOpen)
            {
                serialPort.Close();
                Console.WriteLine("串口关闭成功");
            }
        }

        /// <summary>
        /// 打开串口
        /// </summary>
        /// <param name="duankouhao"></param>
        /// <returns></returns>
        static private bool open_Serial(int duankouhao)
        {
            if (serialPort.IsOpen)
            {
                return true;
            }
            int baud;
            if (!int.TryParse("9600", out baud))
            {
                return false;
            }
            serialPort.PortName = SerialPortTool.GetSerialPortByName(sl[duankouhao]);
            serialPort.BaudRate = baud;
            try
            {
                serialPort.Open();
            }
            catch (System.IO.IOException ioe)
            {
                Console.WriteLine(ioe.Message);
            }
            catch (System.UnauthorizedAccessException ioe)
            {
                Console.WriteLine(ioe.Message);
                return false;
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            if (!serialPort.IsOpen)
            {
                Console.WriteLine(serialPort.PortName + ": 打开串口失败");
                return false;
            }
            Console.WriteLine(serialPort.PortName + ": 打开成功, 速率: " + baud);
            return true;
        }

        public static string HttpPost(string strURL, string strBody)
        {
            Encoding encoding = Encoding.UTF8;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(strURL);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Headers.Add("Authorization:Shangweiji");
            byte[] buffer = encoding.GetBytes(strBody);
            request.ContentLength = buffer.Length;
            request.GetRequestStream().Write(buffer, 0, buffer.Length);
            try
            {
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public static void GetNextOperation(int num)
        {
            string strURL = "http://121.43.107.106:8063/Api/v1/Operation/OpEquipmentSetData";
            OperationNow[num].OperationResult = "OK";
            Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(OperationNow[num])));

            TestResultInfo Test = Processing[num];
            Test.FormerStep = Test.NowStep;
            Test.NowStep = Test.LaterStep;
            var request = (HttpWebRequest)WebRequest.Create("http://121.43.107.106:8063/Api/v1/Operation/MstOperationOrdersGetNextStep?OrderId=" + Test.NowStep);
            request.Headers.Add("Authorization:Shangweiji");
            var response = (HttpWebResponse)request.GetResponse();
            string[] responseString = new StreamReader(response.GetResponseStream()).ReadToEnd().Split(new char[] { '"', ',' });
            Test.LaterStep = responseString[1];
            strURL = "http://121.43.107.106:8063/Api/v1/Result/ResTestResultSetData";
            Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(Test)));

            strURL = "http://121.43.107.106:8063/Api/v1/Operation/OpEquipmentSetData";
            OperationNow[num] = new OperationInfo
            {
                EquipmentId = "Iso_Process",
                OperationTime = DateTime.Now,
                OperationCode = OperationNext[num],
                OperationValue = "1",
                OperationResult = "",
                TerminalIP = "127.0.0.1",
                TerminalName = Dns.GetHostName(),
                revUserId = "SHANGWEIJI"
            };
            OperationNext[num] = responseString[4];
            Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(OperationNow[num])));
        }

        public static void EndOperation(int num)
        {
            string strURL = "http://121.43.107.106:8063/Api/v1/Operation/OpEquipmentSetData";
            OperationNow[num].OperationResult = "OK";
            Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(OperationNow[num])));

            TestResultInfo Test = Processing[num];
            Test.ProcessEnd = DateTime.Now;
            Test.ReStatus = 1;
            Test.FormerStep = "";
            Test.NowStep = "等待：阳性菌加注";
            Test.LaterStep = "";
            strURL = "http://121.43.107.106:8063/Api/v1/Result/ResTestResultSetData";
            Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(Test)));
        }

        public class JsonHelper
        {
            /// <summary>
            /// 将对象序列化为JSON格式
            /// </summary>
            /// <param name="o">对象</param>
            /// <returns>json字符串</returns>
            public static string SerializeObject(object o)
            {
                string json = JsonConvert.SerializeObject(o);
                return json;
            }

            /// <summary>
            /// 解析JSON字符串生成对象实体
            /// </summary>
            /// <typeparam name="T">对象类型</typeparam>
            /// <param name="json">json字符串(eg.{"ID":"112","Name":"石子儿"})</param>
            /// <returns>对象实体</returns>
            public static T DeserializeJsonToObject<T>(string json) where T : class
            {
                JsonSerializer serializer = new JsonSerializer();
                StringReader sr = new StringReader(json);
                object o = serializer.Deserialize(new JsonTextReader(sr), typeof(T));
                T t = o as T;
                return t;
            }

            /// <summary>
            /// 解析JSON数组生成对象实体集合
            /// </summary>
            /// <typeparam name="T">对象类型</typeparam>
            /// <param name="json">json数组字符串(eg.[{"ID":"112","Name":"石子儿"}])</param>
            /// <returns>对象实体集合</returns>
            public static List<T> DeserializeJsonToList<T>(string json) where T : class
            {
                JsonSerializer serializer = new JsonSerializer();
                StringReader sr = new StringReader(json);
                object o = serializer.Deserialize(new JsonTextReader(sr), typeof(List<T>));
                List<T> list = o as List<T>;
                return list;
            }

            /// <summary>
            /// 反序列化JSON到给定的匿名对象.
            /// </summary>
            /// <typeparam name="T">匿名对象类型</typeparam>
            /// <param name="json">json字符串</param>
            /// <param name="anonymousTypeObject">匿名对象</param>
            /// <returns>匿名对象</returns>
            public static T DeserializeAnonymousType<T>(string json, T anonymousTypeObject)
            {
                T t = JsonConvert.DeserializeAnonymousType(json, anonymousTypeObject);
                return t;
            }
        }

        public class TestResultInfo
        {
            public string TestId { get; set; }
            public string ObjectNo { get; set; }
            public string ObjCompany { get; set; }
            public string ObjIncuSeq { get; set; }
            public string TestType { get; set; }
            public string TestStand { get; set; }
            public string TestEquip { get; set; }
            public string TestEquip2 { get; set; }
            public string Description { get; set; }
            public DateTime ProcessStart { get; set; }
            public DateTime ProcessEnd { get; set; }
            public DateTime CollectStart { get; set; }
            public DateTime CollectEnd { get; set; }
            public DateTime TestTime { get; set; }
            public string TestResult { get; set; }
            public string TestPeople { get; set; }
            public string TestPeople2 { get; set; }
            public int ReStatus { get; set; }
            public string RePeople { get; set; }
            public string ReTime { get; set; }
            public string TerminalIP { get; set; }
            public string TerminalName { get; set; }
            public string revUserId { get; set; }
            public string FormerStep { get; set; }
            public string NowStep { get; set; }
            public string LaterStep { get; set; }
        }

        public class OperationInfo
        {
            public string EquipmentId { get; set; }
            public DateTime OperationTime { get; set; }
            public string OperationCode { get; set; }
            public string OperationValue { get; set; }
            public string OperationResult { get; set; }
            public string TerminalIP { get; set; }
            public string TerminalName { get; set; }
            public string revUserId { get; set; }
        }
    }
}
