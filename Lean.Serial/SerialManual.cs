using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Lean.Serial
{

    public partial class SerialManual : Form
    {
        public static String DTAConnectionString = ConfigurationManager.ConnectionStrings["UploadServ"].ToString();
        public SqlConnection DTAConnection = new SqlConnection(DTAConnectionString);
        public Int32 IptSerial, strlen, uploadlen;
        public string IptDate, IptHbnSerial, IptHbn, IptStr, Eptinv, Eptorg, EptDesc, strsn;
        public SerialManual()
        {
            InitializeComponent();
        }
        [DllImport("kernel32.dll")]
        static extern uint GetTickCount();
        //延时函数
        static void Delay(uint ms)
        {
            uint start = GetTickCount();
            while (GetTickCount() - start < ms)
            {
                Application.DoEvents();
            }
        }

        public string TotalTime, upfile, stime, etime;
        protected string getResult = string.Empty;
        private void SerialManual_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Now.AddDays(1 - DateTime.Now.Day).Date;

            dateTimePicker2.Value = DateTime.Now.AddDays(1 - DateTime.Now.Day).Date.AddMonths(1).AddSeconds(-1);
        }
        /// <summary>
        /// 更新品名
        /// </summary>
        /// <param name="gridview">DataGridView控件</param>
        /// <param name="tc_column">是否需要添加列</param>

        protected void UpdateInvMb002()
        {
            GarbageCollect();
            FlushMemory();
            String SalesConnectionString = ConfigurationManager.ConnectionStrings["UploadServ"].ToString();
            SqlConnection SalesConnection = new SqlConnection(SalesConnectionString);
            SalesConnection.Open();
            //DataTable InvTable = new DataTable();
            //string SalesSQL = "SELECT '2300'+LEFT(REPLACE(MB001,' ','')+'                  ',18) +LEFT(LEFT(REPLACE(MB002,' ',''),40)+'                                        ',40) +LEFT(REPLACE(MB006,' ','')+'    ',4)+ LEFT(REPLACE(MB080,' ','')+'                                        ',40)+LEFT(ISNULL(REPLACE(MB003,' ',''),'')+'     ',5)+LEFT(REPLACE(MB110,' ','')+'                                        ',40)  FROM DTA.dbo.INVMB;	";
            string DTASql_GetItemData = "SELECT MB001,REPLACE(MB002,'，',',') FROM DTA.dbo.INVMB WHERE MB002 LIKE '%，%'";
            SqlDataAdapter DTASql_GetItemDa = new SqlDataAdapter(DTASql_GetItemData, SalesConnectionString);
            DataSet DTASql_GetItemDs = new DataSet();
            DTASql_GetItemDa.Fill(DTASql_GetItemDs, "b2F");//SalesData为临时表名可以随意取
            for (int i = 0; i < DTASql_GetItemDs.Tables["b2F"].Rows.Count; i++)
            {
                string DTAUpdateSql = "UPDATE DTA.dbo.INVMB SET MB002='" + DTASql_GetItemDs.Tables["b2F"].Rows[i][1].ToString().TrimEnd() + "' WHERE MB001='" + DTASql_GetItemDs.Tables["b2F"].Rows[i][0].ToString().TrimEnd() + "'";
                SqlCommand DTAUpdateCommand = new SqlCommand(DTAUpdateSql, SalesConnection);
                DTAUpdateCommand.CommandTimeout = 0;
                DTAUpdateCommand.ExecuteNonQuery();

            }
            string TACSql_GetItemData = "SELECT MB001,REPLACE(MB002,'，',',') FROM DTA.dbo.INVMB WHERE MB002 LIKE '%，%'";
            SqlDataAdapter TACSql_GetItemDa = new SqlDataAdapter(TACSql_GetItemData, SalesConnectionString);
            DataSet TACSql_GetItemDs = new DataSet();
            TACSql_GetItemDa.Fill(TACSql_GetItemDs, "b2F");//SalesData为临时表名可以随意取
            for (int i = 0; i < TACSql_GetItemDs.Tables["b2F"].Rows.Count; i++)
            {
                string TACUpdateSql = "UPDATE DTA.dbo.INVMB SET MB002='" + TACSql_GetItemDs.Tables["b2F"].Rows[i][1].ToString().TrimEnd() + "' WHERE MB001='" + TACSql_GetItemDs.Tables["b2F"].Rows[i][0].ToString().TrimEnd() + "'";
                SqlCommand TACUpdateCommand = new SqlCommand(TACUpdateSql, SalesConnection);
                TACUpdateCommand.CommandTimeout = 0;
                TACUpdateCommand.ExecuteNonQuery();

            }
            //InvDa.Fill(InvDs);
            //DgvInventryData.DataSource = InvDs.Tables[0];

            //SqlDataAdapter InvAdapter = new SqlDataAdapter(InvSQL, InvConnection);
            //DataTable Invds = new DataTable(InvSQL);            
            //SalesDa.Fill(SalesDs, "SalesData");//SalesData为临时表名可以随意取

            //Dgvitemtext.DataSource = SalesDs.Tables["SalesData"];
            SalesConnection.Close();
        }


        /// <summary>
        /// DataGridView样式处理
        /// </summary>
        /// <param name="gridview">DataGridView控件</param>
        /// <param name="tc_column">是否需要添加列</param>
        public static void DataGridViewStyle(DataGridView gridview, bool tc_column = true)
        {
            //Form f = new System.Windows.Forms.Form();
            //gridview.AutoGenerateColumns = false;

            //列自动调整大小
            gridview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            //行自动调整大小
            gridview.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //边框样式
            gridview.BorderStyle = BorderStyle.None;
            //背景颜色
            gridview.BackgroundColor = System.Drawing.Color.White;

            //是否停靠父类(全屏)
            gridview.Dock = DockStyle.Fill;

            //列名边框样式
            gridview.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            //鼠标停留控件样式
            //gridview.Cursor = Cursor.Handle;

            gridview.EnableHeadersVisualStyles = false;
            //行名边框样式
            gridview.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            //是否显示行标题的列
            gridview.RowHeadersVisible = false;

            //是否显示添加行的选项
            gridview.AllowUserToAddRows = false;
            //是否显示删除行的选项
            gridview.AllowUserToDeleteRows = false;
            //设置标题列高度是否为可调或自动
            gridview.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            //设置用户是否可以多选
            gridview.MultiSelect = false;
            //设置用户对控件内容是否为只读
            gridview.ReadOnly = true;
            //设置行标头宽度
            gridview.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            //设置选择单元模式
            gridview.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            //标题文字居中
            gridview.ColumnHeadersDefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;

            //奇数行变色
            gridview.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(234, 234, 234);

            if (tc_column)
            {
                //设置默认填充列自动填充
                DataGridViewTextBoxColumn TextBoxColumn = new DataGridViewTextBoxColumn();
                TextBoxColumn.Name = "";
                TextBoxColumn.HeaderText = "";

                //gridview.Columns.Insert(gridview.Columns.Count, TextBoxColumn);
                gridview.Columns.Insert(gridview.ColumnCount, TextBoxColumn);
                TextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
                //取消该列排序
                TextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            }

            //点击立即进入编辑模式
            gridview.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void button1_Click(object sender, EventArgs e)
        {



            upfile = "sndta.txt";
            //记时器
            TimeSpan ts1 = new TimeSpan(DateTime.Now.Ticks); //获取当前时间的刻度数

            //延时20分钟'1200000
            Delay(0);
            //更新品名信息
            //UpdateInvMb002();
            //传递数据到DataGridView


            DateTime updatstime = dateTimePicker1.Value; //DateTime.Now.AddDays(-1);//
            DateTime updatetime = dateTimePicker2.Value; //DateTime.Now.AddDays(-1);//
                                                         //etime = sdate + "000000";
                                                         //etime = edate + "230000";
            stime = updatstime.ToString("yyyyMMdd") + "000000";
            //etime = updatetime.ToString("yyyyMMdd") + "235959";
            etime = updatetime.ToString("yyyyMMdd") + "235959";

            GetSerialData();





            //文件处理
            //文件路径
            //Application.StartupPath + "\\BackFile\\"
            File.Delete(@Application.StartupPath + "\\serialTxt\\" + upfile);
            //File.Delete(@"d:\DSI\TCJ\sndta.xls");

            //File.Delete(@"d:\DSI\*.*");
            //删除文件
            //string dt = DateTime.Now.ToString("yyyy-MM-dd");
            //ExportToTxt(DgvSalesData, dt + "SalesData.txt");
            //ExportToTxt(DgvInventryData, dt + "InventryData.txt");


            //ExportToTxt(Dgvitemtext, dt + "ItemText.txt");
            //ExportToTxt(Dgvsuppdata, dt + "SupplierName.txt");
            //ExportToTxt(DgvModelNumbe, dt + "ModelNumber.txt");

            //ExportToTxt(DgvGetInvStockData, dt + "InvStockData.txt");
            //ExportToTxt(DgvGetItemData, dt + "ItemData.txt");
            //ExportToTxt(DgvGetPriceStdData, dt + "PriceStdData.txt");
            //ExportToTxt(DgvGetSupplierData, dt + "SupplierData.txt");
            //ExportToTxt(DgvGetPurchaseData, dt + "PurchaseData.txt");

            //ExportToTxt(DgvGetInStorageData, dt + "InStorageData.txt");
            //ExportToTxt(DgvGetOutStorageData, dt + "OutStorageData.txt");
            //ExportToTxt(DgvGetResidualData, dt + "ResidualData.txt");
            //ExportToTxt(DgvGetInvStockData, dt + "ReserveData.txt");
            ////文件处理


            //保存到本地

            FtpExportToTxt(DgvSndata, upfile);
            //DGVToExcel(DgvSnDataCsv, "sndta.xls");

            //保存到本地待上传FTP
            string redt = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");

            string APP_Path = Application.StartupPath;//获取启动了应用程序的可执行文件的路径，不包括可执行文件的名称            
            string XmlFile = APP_Path + "\\SerialConfig.xml";


            string uid = Helper_Xml.Read(XmlFile, "/Root/FTP/UserInfo", "Uid").ToString();
            string pwd = Helper_Xml.Read(XmlFile, "/Root/FTP/UserInfo", "Pwd").ToString();
            string ftpserv = Helper_Xml.Read(XmlFile, "/Root/FTP/ServInfo", "ServIpaddr").ToString();

            Helper_Ftp Helper_Ftp = new Helper_Ftp(ftpserv, uid, pwd);


            if (Helper_Ftp.fileCheckExist(ftpserv, upfile) == true)
            {
                Helper_Ftp.DeleteFileName(upfile);
                //Helper_Ftp.Rename("InventryData.txt", redt + "InventryData.txt");
            }

            //if (Helper_Ftp.fileCheckExist("ftp2.teac.co.jp", "sndta.xls") == true)
            //{
            //    Helper_Ftp.DeleteFileName("sndta.xls");
            //    //Helper_Ftp.Rename("InventryData.txt", redt + "InventryData.txt");
            //}

            if (DgvSndata.Rows.Count != 0)
            {
                Helper_Ftp.Upload(@Application.StartupPath + "\\serialTxt\\" + upfile);
                //Helper_Ftp.Upload(@"d:\DSI\TCJ\sndta.xls");


                //Helper_Ftp.Upload(@"d:\DSI\TCJ\ReserveData.txt");

                //FtpStatusCode InventryData = UploadFun(@"d:\DSI\TCJ\InventryData.txt", "ftp://172.20.233.40/InventryData.txt");
                //FtpStatusCode SalesData = UploadFun(@"d:\DSI\TCJ\SalesData.txt", "ftp://172.20.233.40/SalesData.txt");

                //FtpStatusCode InvStockData = UploadFun(@"d:\DSI\TCJ\InvStockData.txt", "ftp://172.20.233.40/InvStockData.txt");
                //FtpStatusCode ItemData = UploadFun(@"d:\DSI\TCJ\ItemData.txt", "ftp://172.20.233.40/ItemData.txt");
                //FtpStatusCode PriceStdData = UploadFun(@"d:\DSI\TCJ\PriceStdData.txt", "ftp://172.20.233.40/PriceStdData.txt");
                //FtpStatusCode SupplierData = UploadFun(@"d:\DSI\TCJ\SupplierData.txt", "ftp://172.20.233.40/SupplierData.txt");
                //FtpStatusCode PurchaseData = UploadFun(@"d:\DSI\TCJ\PurchaseData.txt", "ftp://172.20.233.40/PurchaseData.txt");

                //上传到FTP文件服务器


            }
            //显示运行时间
            TimeSpan ts2 = new TimeSpan(DateTime.Now.Ticks);
            TimeSpan ts = ts2.Subtract(ts1).Duration(); //时间差的绝对值
            string spanTotalSeconds = ts.TotalSeconds.ToString(); //执行时间的总秒数
            string spanTime = "FTP Uplaod,共花费" + ts.Hours.ToString() + "小时" + ts.Minutes.ToString() + "分" + ts.Seconds.ToString() + "秒！"; //以X小时X分X秒的格式现实执行时间
            TotalTime = ts.Hours.ToString() + "小时" + ts.Minutes.ToString() + "分" + ts.Seconds.ToString() + "秒！";
            //显示运行时间
            //UpdateTime(DateTime.Now.ToString(), spanTime);//当前时间
            //文件处理
            uploadlen = DgvSndata.Rows.Count;
            SendMail();
            //关闭窗体
            this.Close();
            MessageBox.Show(updatetime.ToString("yyyy-MM-dd") + "的数据上传完成！");

        }

        public void Eptsave()
        {
            if (DTAConnection.State == ConnectionState.Closed)
            {
                DTAConnection.Open();
            }

            SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_SCANNER_UP] ([OUTS001],[OUTS002],[OUTS003],[OUTS004],[OUTS005],[OUTS006],[OUTS007],[OUTS008],[OUTSYSUID],[OUTHNAME],[OUTHIP],[OUTHMAC],[OUTUSER],[OUTDTIME],[studfchar1],[studfchar2],[studfchar3],[studfchar4],[studfint1],[studfint2],[studfint3],[studfint4])VALUES  ('" + IptDate + "','" + this.Eptinv + "','" + this.Eptorg + "', '" + this.IptHbnSerial + "','" + this.IptHbn + "','" + this.IptStr + "','" + this.EptDesc + "',1,'OutStock','" + Helper_Hard.GetComputerName() + "','" + Helper_Hard.GetIPAddress() + "','" + Helper_Hard.GetMacAddress() + "','USER','" + DateTime.Now.ToString("yyyyMMddHHmmss") + "','','','','','0','0','0','0') ", DTAConnection);
            InsproaSqlcom.ExecuteNonQuery();

            if (DTAConnection.State == ConnectionState.Open)
            {
                DTAConnection.Close();
            }

        }
        protected void GetSerialData() //为datagridview綁定数据源
        {


            //记时器
            TimeSpan ts1 = new TimeSpan(DateTime.Now.Ticks); //获取当前时间的刻度数

            //延时20分钟1200000
            Delay(0);
            DataTable DTAStockTable = new DataTable();
            string DTAStockSQL = "DELETE[LeanSerial].[dbo].[DTASSET_SCANNER_UP] WHERE OUTS001>= '" + stime + "' AND [OUTS001]<= '" + etime + "';SELECT [OUTSERIAL], [OUTINVOICE],[OUTDESC],[OUTDATE],[OUTQTY],[OGION]  FROM [LeanSerial].[dbo].[DTASSET_SCANNER_ALL] WHERE [OUTDTIME]>= '" + stime + "' AND [OUTDTIME]<= '" + etime + "';";
            SqlDataAdapter DTAStockAdapter = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
            DTAStockAdapter.Fill(DTAStockTable);
            //显示品名
            if (DTAStockTable.Rows.Count >= 1)
            {

                for (int tbi = 1; tbi <= DTAStockTable.Rows.Count; tbi++)
                {
                    if (DTAStockTable.Rows[tbi - 1][0].ToString().Length != 0)
                    {
                        int fs = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Length);
                        int tsss = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(fs - 2, 2));

                        IptSerial = tsss;
                        IptHbnSerial = DTAStockTable.Rows[tbi - 1][0].ToString();
                        IptHbn = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(0, 10);
                        IptDate = DTAStockTable.Rows[tbi - 1][3].ToString();
                        Eptinv = DTAStockTable.Rows[tbi - 1][1].ToString();
                        Eptorg = DTAStockTable.Rows[tbi - 1][5].ToString();
                        EptDesc = DTAStockTable.Rows[tbi - 1][2].ToString();

                        if (IptSerial == 0)
                        {
                            IptSerial = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(fs - 3, 3));
                            for (int i = 0; i < IptSerial; i++)
                            {
                                string strA = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3);
                                Regex reg = new Regex("^[0-9]+$");
                                Match ma = reg.Match(strA);
                                /// <summary>
                                /// //判断是数字字串还是字符字串
                                /// </summary>

                                if (ma.Success)
                                {
                                    int len = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3)) + i;
                                    if (Convert.ToInt32(len.ToString().Length) != 7)
                                    {
                                        IptStr = String.Format("{0:D7}", len);
                                    }
                                    else
                                    {
                                        IptStr = len.ToString();
                                    }
                                    Eptsave();
                                }
                                else
                                {
                                    //判断字母出现的位置
                                    //int f = 0;
                                    //foreach (char c in DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7))
                                    //{
                                    //    f++;
                                    //    if (char.IsLetter(c))
                                    //    {
                                    //        //int iii = ii - 1;
                                    //        string ss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7).Substring(f, 7 - f);
                                    //        string sss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7).Substring(0, f);
                                    //    }
                                    //}
                                    char cc = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3).Last((c) => { return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'); });
                                    int index = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3).IndexOf(cc) + 1;
                                    string sss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3).Substring(0, index);
                                    string ss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3).Substring(index, fs - 10 - 3 - index);

                                    int len = Convert.ToInt32(ss) + i;
                                    int lena = Convert.ToInt32(ss.Length.ToString());
                                    if (Convert.ToInt32(len.ToString().Length) != lena)
                                    {
                                        string d = "{0:D" + lena + "}";
                                        IptStr = String.Format(d, len);
                                        //str = len.ToString(d);
                                    }
                                    else
                                    {
                                        IptStr = len.ToString();
                                    }
                                    //判断字母出现的位置

                                    IptStr = sss + Convert.ToString(IptStr);
                                    /// <summary>
                                    /// //判断是数字字串还是字符字串
                                    /// </summary>
                                    Eptsave();
                                }
                            }
                        }


                        else
                        {
                            IptSerial = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(fs - 2, 2));
                            for (int i = 0; i < IptSerial; i++)
                            {
                                string strA = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2);
                                Regex reg = new Regex("^[0-9]+$");
                                Match ma = reg.Match(strA);
                                /// <summary>
                                /// //判断是数字字串还是字符字串
                                /// </summary>

                                if (ma.Success)
                                {
                                    int len = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2)) + i;
                                    if (Convert.ToInt32(len.ToString().Length) != 7)
                                    {
                                        IptStr = String.Format("{0:D7}", len);
                                    }
                                    else
                                    {
                                        IptStr = len.ToString();
                                    }
                                    Eptsave();
                                }
                                else
                                {
                                    //判断字母出现的位置
                                    //int f = 0;
                                    //foreach (char c in DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7))
                                    //{
                                    //    f++;
                                    //    if (char.IsLetter(c))
                                    //    {
                                    //        //int iii = ii - 1;
                                    //        string ss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7).Substring(f, 7 - f);
                                    //        string sss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7).Substring(0, f);
                                    //    }
                                    //}
                                    char cc = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2).Last((c) => { return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'); });
                                    int index = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2).LastIndexOf(cc) + 1;
                                    string sss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2).Substring(0, index);
                                    string ss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2).Substring(index, fs - 10 - 2 - index);

                                    int len = Convert.ToInt32(ss) + i;
                                    int lena = Convert.ToInt32(ss.Length.ToString());
                                    if (Convert.ToInt32(len.ToString().Length) != lena)
                                    {
                                        string d = "{0:D" + lena + "}";
                                        IptStr = String.Format(d, len);
                                        //str = len.ToString(d);
                                    }
                                    else
                                    {
                                        IptStr = len.ToString();
                                    }
                                    //判断字母出现的位置

                                    IptStr = sss + Convert.ToString(IptStr);
                                    /// <summary>
                                    /// //判断是数字字串还是字符字串
                                    /// </summary>
                                    Eptsave();
                                }
                            }
                        }

                    }
                }










                GarbageCollect();
                FlushMemory();
                String InvConnectionString = ConfigurationManager.ConnectionStrings["UploadServ"].ToString();
                SqlConnection InvConnection = new SqlConnection(InvConnectionString);
                InvConnection.Open();
                //DataTable InvTable = new DataTable();
                //            string InvSQL = "SELECT   Category, MaterialCode, MaterialText, MaterialCode2, StorageLocation, Qty , convert(float,Amount) AS Amount, MaterialType  "+
                //" FROM "+
                //" (SELECT rtrim(cast(2400 AS varchar(255))+cast(Year(getdate())*10000+Month(getdate())*100+Day(getdate()) AS varchar(255))+CAST( isnull(TACHK.dbo.BOMMC.MC010,'') as varchar(255))) AS Category, rtrim(TACHK.dbo.INVMB.MB001) AS MaterialCode, LEFT(TACHK.dbo.INVMB.MB002,30) AS MaterialText, rtrim(TACHK.dbo.INVMB.MB001) AS MaterialCode2, 'HKD2400'+CAST( CASE WHEN MB005='6001' THEN 'ALPS' ELSE 'OnTheWay' END  as varchar(255))StorageLocation, 'H'+cast(TACHK.dbo.INVMB.MB064 as varchar(255)) AS Qty, TACHK.dbo.INVMB.MB065 AS Amount, CASE WHEN MB005='6001' THEN 'ROH' WHEN MB005 ='6002' THEN 'HALB' WHEN MB005 ='6003' THEN 'FERT' ELSE '' END MaterialType "+
                //" FROM TACHK.dbo.BOMMC RIGHT JOIN TACHK.dbo.INVMB ON TACHK.dbo.BOMMC.MC001 = TACHK.dbo.INVMB.MB001 "+
                //" WHERE (((TACHK.dbo.INVMB.MB064)>0) AND ((TACHK.dbo.INVMB.MB005)<>'6004'))) AS TAC "+
                //" UNION ALL SELECT * FROM "+
                //" (SELECT rtrim(cast(2300 as varchar(255))+cast(Year(getdate())*10000+Month(getdate())*100+Day(getdate()) as varchar(255))+cast(isnull(DTA.dbo.BOMMC.MC010,'') as varchar(255))) AS Category, rtrim(DTA.dbo.INVMB.MB001) AS MaterialCode, LEFT(DTA.dbo.INVMB.MB002,30) AS MaterialText, rtrim(DTA.dbo.INVMB.MB001) AS MaterialCode2, 'HKD2300Dongguan' AS StorageLocation, 'H' +CAST( DTA.dbo.INVMB.MB064 as varchar(255)) AS Qty, DTA.dbo.INVMB.MB065/B.MG003 AS Amount, CASE WHEN MB005='6001' THEN 'ROH' WHEN MB005 ='6002' THEN 'HALB' WHEN MB005 ='6003' THEN 'FERT' ELSE '' END MaterialType "+
                //" FROM DTA.dbo.BOMMC RIGHT JOIN DTA.dbo.INVMB ON DTA.dbo.BOMMC.MC001 = DTA.dbo.INVMB.MB001, "+
                //" (SELECT DTA.dbo.CMSMG.MG001, DTA.dbo.CMSMG.MG003 FROM DTA.dbo.CMSMG INNER JOIN (SELECT MG001, Max(MG002) AS mMG002 FROM DTA.dbo.CMSMG GROUP BY MG001 HAVING DTA.dbo.CMSMG.MG001='HKD') AS h ON DTA.dbo.CMSMG.MG001 = h.MG001 WHERE DTA.dbo.CMSMG.MG002=h.mMG002) AS B "+
                //" WHERE (((DTA.dbo.INVMB.MB064)>0) AND ((DTA.dbo.INVMB.MB005)<'6003'))) AS DTApart "+
                //" UNION ALL SELECT * FROM  "+
                //" (SELECT rtrim(cast(2300 as varchar(255))+CAST(Year(getdate())*10000+Month(getdate())*100+Day(getdate()) as varchar(255))+CAST(DTA.dbo.BOMMC.MC010 as varchar(255))) AS Category, rtrim(DTA.dbo.INVMB.MB001) AS MaterialCode, LEFT(DTA.dbo.INVMB.MB002,30) AS MaterialText, rtrim(DTA.dbo.INVMB.MB001) AS MaterialCode2, 'HKD2300Dongguan' AS StorageLocation, 'H'+cast(StockDataB.MC007SUM as varchar(255)) AS Qty, StockDataB.MC008SUM/B.MG003 AS Amount, CASE WHEN MB005='6001' THEN 'ROH' WHEN MB005 ='6002' THEN 'HALB' WHEN MB005 ='6003' THEN 'FERT' ELSE '' END MaterialType "+
                //" FROM (DTA.dbo.BOMMC RIGHT JOIN DTA.dbo.INVMB ON DTA.dbo.BOMMC.MC001 = DTA.dbo.INVMB.MB001) INNER JOIN  "+
                //" (SELECT DTA.dbo.INVMB.MB001, Sum(DTA.dbo.INVMC.MC007) AS MC007SUM, Sum(DTA.dbo.INVMC.MC008) AS MC008SUM FROM DTA.dbo.INVMB LEFT JOIN DTA.dbo.INVMC ON DTA.dbo.INVMB.MB001 = DTA.dbo.INVMC.MC001 WHERE (((DTA.dbo.INVMB.MB005)='6003') AND ((DTA.dbo.INVMC.MC002)<>'H')) GROUP BY DTA.dbo.INVMB.MB001 HAVING (((Sum(DTA.dbo.INVMC.MC007))>0))) AS StockDataB  "+
                //" ON DTA.dbo.INVMB.MB001 = StockDataB.MB001,  "+
                //" (SELECT DTA.dbo.CMSMG.MG001, DTA.dbo.CMSMG.MG003 FROM DTA.dbo.CMSMG INNER JOIN (SELECT MG001, Max(MG002) AS mMG002 FROM DTA.dbo.CMSMG GROUP BY MG001 HAVING DTA.dbo.CMSMG.MG001='HKD') AS h ON DTA.dbo.CMSMG.MG001 = h.MG001 WHERE DTA.dbo.CMSMG.MG002=h.mMG002) AS B "+
                //" WHERE ((DTA.dbo.INVMB.MB005)='6003')) AS InventryDTAB "+
                //" UNION ALL SELECT * FROM  "+
                //" (SELECT rtrim(cast(2300 as varchar(255))+cast(Year(getdate())*10000+Month(getdate())*100+Day(getdate()) as varchar(255))+cast(DTA.dbo.BOMMC.MC010 AS varchar(255))) AS Category, rtrim(DTA.dbo.INVMB.MB001) AS MaterialCode, LEFT(DTA.dbo.INVMB.MB002,30) AS MaterialText, rtrim(DTA.dbo.INVMB.MB001) AS MaterialCode2, 'HKD2300BFR' AS StorageLocation, 'H'+CAST(StockDataH.MC007 as varchar(255)) AS Qty,StockDataH.MC008/B.MG003 AS Amount, CASE WHEN MB005='6001' THEN 'ROH' WHEN MB005 ='6002' THEN 'HALB' WHEN MB005 ='6003' THEN 'FERT' ELSE '' END MaterialType "+
                //" FROM (DTA.dbo.BOMMC RIGHT JOIN DTA.dbo.INVMB ON DTA.dbo.BOMMC.MC001 = DTA.dbo.INVMB.MB001) INNER JOIN  "+
                //" (SELECT DTA.dbo.INVMB.MB001, DTA.dbo.INVMC.MC007, DTA.dbo.INVMC.MC008 FROM DTA.dbo.INVMB LEFT JOIN DTA.dbo.INVMC ON DTA.dbo.INVMB.MB001 = DTA.dbo.INVMC.MC001 WHERE (((DTA.dbo.INVMB.MB005)='6003') AND ((DTA.dbo.INVMC.MC002)='H'))) AS StockDataH "+
                //" ON DTA.dbo.INVMB.MB001 = StockDataH.MB001, "+
                //" (SELECT DTA.dbo.CMSMG.MG001, DTA.dbo.CMSMG.MG003 FROM DTA.dbo.CMSMG INNER JOIN (SELECT MG001, Max(MG002) AS mMG002 FROM DTA.dbo.CMSMG GROUP BY MG001 HAVING DTA.dbo.CMSMG.MG001='HKD') AS h ON DTA.dbo.CMSMG.MG001 = h.MG001 WHERE DTA.dbo.CMSMG.MG002=h.mMG002) AS B "+
                //" WHERE ((DTA.dbo.INVMB.MB005)='6003')) AS InventryDTAH";
                string InvSQL = "SELECT  日付 + 出荷伝票番号 + 明細番号 + 品目コード + 合計数量 + シリアルNO + 箱の入り数 + 仕入先 + 品番テキスト as sndata from( " +
                                " SELECT left(ltrim([OUTS001]) + '        ', 8) as 日付, left(ltrim(OUTS002) + '         ', 9) as 出荷伝票番号, " +
                                " right(OUTS006, 3) 明細番号, left(ltrim([OUTS005]) + '          ', 10) as 品目コード,  RIGHT('0000000' + CAST(case when right(OUTS004, 2) = '' then 1 " +
                                  " when right(OUTS004, 2) = 1 then 1  " +

                                  " when right(OUTS004, 2) = 2 then 2  " +

                                  " when right(OUTS004, 2) = 3 then 3  " +

                                  " when right(OUTS004, 2) = 4 then 4 " +

                                  " when right(OUTS004, 2) = 5 then 5 " +

                                  " when right(OUTS004, 2) = 6 then 6 " +

                                  " when right(OUTS004, 2) = 7 then 7 " +

                                  " when right(OUTS004, 2) = 8 then 8 " +

                                  " when right(OUTS004, 2) = 9 then 9 " +

                                  " when right(OUTS004, 2) = 10 then 10 " +

                                "   when right(OUTS004, 2) = 0 then 100 " +
                                 "  else 1 end AS NVARCHAR(50)), 7) 合計数量, RIGHT('0000000' + CAST(OUTS006  AS nvarchar(50)), 7)  シリアルNO,  " +
                                " '001' as 箱の入り数,left(ltrim(OUTS007) + '                    ', 20) 仕入先,left(REPLACE(MB002,' ','') + '                                        ', 40)品番テキスト" +
                                " FROM[LeanSerial].[dbo].[DTASSET_SCANNER_UP] LEFT JOIN[LeanSerial].[dbo].[DTASSET_INVMB] ON MB001 = OUTS005 " +
                                " where [OUTDTIME]>='" + stime + "' and [OUTDTIME]<='" + etime + "' and OUTS006<>'0000000' ) as bb order by 日付;";


                SqlDataAdapter InvDa = new SqlDataAdapter(InvSQL, InvConnectionString);
                DataSet InvDs = new DataSet();
                //InvDa.Fill(InvDs);
                //DgvInventryData.DataSource = InvDs.Tables[0];

                //SqlDataAdapter InvAdapter = new SqlDataAdapter(InvSQL, InvConnection);
                //DataTable Invds = new DataTable(InvSQL);            
                InvDa.Fill(InvDs, "InvData");//InvData为临时表名可以随意取


                if (InvDs == null || InvDs.Tables.Count == 0 || InvDs.Tables["InvData"].Rows.Count == 0)
                {


                    return;
                }
                else
                {


                    DgvSndata.DataSource = InvDs.Tables["InvData"];
                }

                InvConnection.Close();

            }

            else
            {

                return;
            }
        }
        protected void GetSerialData24() //为datagridview綁定数据源
        {
            GarbageCollect();
            FlushMemory();
            String InvConnectionString = ConfigurationManager.ConnectionStrings["UploadServ"].ToString();
            SqlConnection InvConnection = new SqlConnection(InvConnectionString);
            InvConnection.Open();
            //DataTable InvTable = new DataTable();
            //            string InvSQL = "SELECT   Category, MaterialCode, MaterialText, MaterialCode2, StorageLocation, Qty , convert(float,Amount) AS Amount, MaterialType  "+
            //" FROM "+
            //" (SELECT rtrim(cast(2400 AS varchar(255))+cast(Year(getdate())*10000+Month(getdate())*100+Day(getdate()) AS varchar(255))+CAST( isnull(TACHK.dbo.BOMMC.MC010,'') as varchar(255))) AS Category, rtrim(TACHK.dbo.INVMB.MB001) AS MaterialCode, LEFT(TACHK.dbo.INVMB.MB002,30) AS MaterialText, rtrim(TACHK.dbo.INVMB.MB001) AS MaterialCode2, 'HKD2400'+CAST( CASE WHEN MB005='6001' THEN 'ALPS' ELSE 'OnTheWay' END  as varchar(255))StorageLocation, 'H'+cast(TACHK.dbo.INVMB.MB064 as varchar(255)) AS Qty, TACHK.dbo.INVMB.MB065 AS Amount, CASE WHEN MB005='6001' THEN 'ROH' WHEN MB005 ='6002' THEN 'HALB' WHEN MB005 ='6003' THEN 'FERT' ELSE '' END MaterialType "+
            //" FROM TACHK.dbo.BOMMC RIGHT JOIN TACHK.dbo.INVMB ON TACHK.dbo.BOMMC.MC001 = TACHK.dbo.INVMB.MB001 "+
            //" WHERE (((TACHK.dbo.INVMB.MB064)>0) AND ((TACHK.dbo.INVMB.MB005)<>'6004'))) AS TAC "+
            //" UNION ALL SELECT * FROM "+
            //" (SELECT rtrim(cast(2300 as varchar(255))+cast(Year(getdate())*10000+Month(getdate())*100+Day(getdate()) as varchar(255))+cast(isnull(DTA.dbo.BOMMC.MC010,'') as varchar(255))) AS Category, rtrim(DTA.dbo.INVMB.MB001) AS MaterialCode, LEFT(DTA.dbo.INVMB.MB002,30) AS MaterialText, rtrim(DTA.dbo.INVMB.MB001) AS MaterialCode2, 'HKD2300Dongguan' AS StorageLocation, 'H' +CAST( DTA.dbo.INVMB.MB064 as varchar(255)) AS Qty, DTA.dbo.INVMB.MB065/B.MG003 AS Amount, CASE WHEN MB005='6001' THEN 'ROH' WHEN MB005 ='6002' THEN 'HALB' WHEN MB005 ='6003' THEN 'FERT' ELSE '' END MaterialType "+
            //" FROM DTA.dbo.BOMMC RIGHT JOIN DTA.dbo.INVMB ON DTA.dbo.BOMMC.MC001 = DTA.dbo.INVMB.MB001, "+
            //" (SELECT DTA.dbo.CMSMG.MG001, DTA.dbo.CMSMG.MG003 FROM DTA.dbo.CMSMG INNER JOIN (SELECT MG001, Max(MG002) AS mMG002 FROM DTA.dbo.CMSMG GROUP BY MG001 HAVING DTA.dbo.CMSMG.MG001='HKD') AS h ON DTA.dbo.CMSMG.MG001 = h.MG001 WHERE DTA.dbo.CMSMG.MG002=h.mMG002) AS B "+
            //" WHERE (((DTA.dbo.INVMB.MB064)>0) AND ((DTA.dbo.INVMB.MB005)<'6003'))) AS DTApart "+
            //" UNION ALL SELECT * FROM  "+
            //" (SELECT rtrim(cast(2300 as varchar(255))+CAST(Year(getdate())*10000+Month(getdate())*100+Day(getdate()) as varchar(255))+CAST(DTA.dbo.BOMMC.MC010 as varchar(255))) AS Category, rtrim(DTA.dbo.INVMB.MB001) AS MaterialCode, LEFT(DTA.dbo.INVMB.MB002,30) AS MaterialText, rtrim(DTA.dbo.INVMB.MB001) AS MaterialCode2, 'HKD2300Dongguan' AS StorageLocation, 'H'+cast(StockDataB.MC007SUM as varchar(255)) AS Qty, StockDataB.MC008SUM/B.MG003 AS Amount, CASE WHEN MB005='6001' THEN 'ROH' WHEN MB005 ='6002' THEN 'HALB' WHEN MB005 ='6003' THEN 'FERT' ELSE '' END MaterialType "+
            //" FROM (DTA.dbo.BOMMC RIGHT JOIN DTA.dbo.INVMB ON DTA.dbo.BOMMC.MC001 = DTA.dbo.INVMB.MB001) INNER JOIN  "+
            //" (SELECT DTA.dbo.INVMB.MB001, Sum(DTA.dbo.INVMC.MC007) AS MC007SUM, Sum(DTA.dbo.INVMC.MC008) AS MC008SUM FROM DTA.dbo.INVMB LEFT JOIN DTA.dbo.INVMC ON DTA.dbo.INVMB.MB001 = DTA.dbo.INVMC.MC001 WHERE (((DTA.dbo.INVMB.MB005)='6003') AND ((DTA.dbo.INVMC.MC002)<>'H')) GROUP BY DTA.dbo.INVMB.MB001 HAVING (((Sum(DTA.dbo.INVMC.MC007))>0))) AS StockDataB  "+
            //" ON DTA.dbo.INVMB.MB001 = StockDataB.MB001,  "+
            //" (SELECT DTA.dbo.CMSMG.MG001, DTA.dbo.CMSMG.MG003 FROM DTA.dbo.CMSMG INNER JOIN (SELECT MG001, Max(MG002) AS mMG002 FROM DTA.dbo.CMSMG GROUP BY MG001 HAVING DTA.dbo.CMSMG.MG001='HKD') AS h ON DTA.dbo.CMSMG.MG001 = h.MG001 WHERE DTA.dbo.CMSMG.MG002=h.mMG002) AS B "+
            //" WHERE ((DTA.dbo.INVMB.MB005)='6003')) AS InventryDTAB "+
            //" UNION ALL SELECT * FROM  "+
            //" (SELECT rtrim(cast(2300 as varchar(255))+cast(Year(getdate())*10000+Month(getdate())*100+Day(getdate()) as varchar(255))+cast(DTA.dbo.BOMMC.MC010 AS varchar(255))) AS Category, rtrim(DTA.dbo.INVMB.MB001) AS MaterialCode, LEFT(DTA.dbo.INVMB.MB002,30) AS MaterialText, rtrim(DTA.dbo.INVMB.MB001) AS MaterialCode2, 'HKD2300BFR' AS StorageLocation, 'H'+CAST(StockDataH.MC007 as varchar(255)) AS Qty,StockDataH.MC008/B.MG003 AS Amount, CASE WHEN MB005='6001' THEN 'ROH' WHEN MB005 ='6002' THEN 'HALB' WHEN MB005 ='6003' THEN 'FERT' ELSE '' END MaterialType "+
            //" FROM (DTA.dbo.BOMMC RIGHT JOIN DTA.dbo.INVMB ON DTA.dbo.BOMMC.MC001 = DTA.dbo.INVMB.MB001) INNER JOIN  "+
            //" (SELECT DTA.dbo.INVMB.MB001, DTA.dbo.INVMC.MC007, DTA.dbo.INVMC.MC008 FROM DTA.dbo.INVMB LEFT JOIN DTA.dbo.INVMC ON DTA.dbo.INVMB.MB001 = DTA.dbo.INVMC.MC001 WHERE (((DTA.dbo.INVMB.MB005)='6003') AND ((DTA.dbo.INVMC.MC002)='H'))) AS StockDataH "+
            //" ON DTA.dbo.INVMB.MB001 = StockDataH.MB001, "+
            //" (SELECT DTA.dbo.CMSMG.MG001, DTA.dbo.CMSMG.MG003 FROM DTA.dbo.CMSMG INNER JOIN (SELECT MG001, Max(MG002) AS mMG002 FROM DTA.dbo.CMSMG GROUP BY MG001 HAVING DTA.dbo.CMSMG.MG001='HKD') AS h ON DTA.dbo.CMSMG.MG001 = h.MG001 WHERE DTA.dbo.CMSMG.MG002=h.mMG002) AS B "+
            //" WHERE ((DTA.dbo.INVMB.MB005)='6003')) AS InventryDTAH";
            string InvSQL = "SELECT left([budate],8) as 日付, RIGHT('000000000' + CAST( [buno]  AS nvarchar(50)),9) as 出荷伝票番号, " +
                            " right([buprosn],3) 明細番号, [buitem] as 品目コード, '0000001' as 合計数量, [buprosn] シリアルNO, case when right([bumocta],2)='' then '001' else '0'+ right([bumocta],2) end 箱の入り数 " +
                            " FROM[LeanSerial].[dbo].[DTASSET_BUKAN] where left(budate,6)=(select CONVERT (nvarchar(6),GETDATE(),112)) and len([buprosn])+len([buitem])=17 and buprosn<>'0000000'  order by 日付";



            SqlDataAdapter InvDa = new SqlDataAdapter(InvSQL, InvConnectionString);
            DataSet InvDs = new DataSet();
            //InvDa.Fill(InvDs);
            //DgvInventryData.DataSource = InvDs.Tables[0];

            //SqlDataAdapter InvAdapter = new SqlDataAdapter(InvSQL, InvConnection);
            //DataTable Invds = new DataTable(InvSQL);            
            InvDa.Fill(InvDs, "InvData");//InvData为临时表名可以随意取
            DgvSnDataCsv.DataSource = InvDs.Tables["InvData"];
            InvConnection.Close();

        }

        //private void BtnOutTxt_Click(object sender, EventArgs e)
        //{

        //    File.Delete(@"d:\DSI\TCJ\InventryData.txt");
        //    File.Delete(@"d:\DSI\TCJ\SalesData.txt");
        //    //删除文件
        //    string dt = DateTime.Now.ToString("yyyy-MM-dd");
        //    ExportToTxt(DgvSalesData, dt + "SalesData.txt");
        //    ExportToTxt(DgvInventryData, dt + "InventryData.txt");
        //    //保存到本地
        //    FtpExportToTxt(DgvSalesData, "SalesData.txt");
        //    FtpExportToTxt(DgvInventryData, "InventryData.txt");
        //    //保存到本地待上传FTP
        //    FtpStatusCode InventryData = UploadFun(@"d:\DSI\TCJ\InventryData.txt", "ftp://172.20.233.40/InventryData.txt");
        //    FtpStatusCode SalesData = UploadFun(@"d:\DSI\TCJ\SalesData.txt", "ftp://172.20.233.40/SalesData.txt");
        //    //上传到FTP文件服务器

        //}



        public void ExportToTxt(DataGridView DGVDTA, string Fn)//导出到本地硬盘
        {
            GarbageCollect();
            FlushMemory();

            string Filename;

            Filename = @Application.StartupPath + "\\serialTxt\\" + Fn.ToString();
            StreamWriter sw = new StreamWriter(Filename, false, System.Text.Encoding.GetEncoding("GB2312"));

            try
            {
                int[] colContentLength = new int[DGVDTA.ColumnCount];//控制列宽
                for (int col = 0; col < DGVDTA.ColumnCount; col++)
                {
                    colContentLength[col] = DGVDTA.Columns[col].HeaderText.Length;
                }

                for (int row = 0; row < DGVDTA.RowCount - 1; row++)
                {
                    for (int col = 0; col < DGVDTA.ColumnCount; col++)
                    {
                        if (DGVDTA.Rows[row].Cells[col].Value != null)
                        {
                            if (DGVDTA.Rows[row].Cells[col].Value.ToString().Length > colContentLength[col])
                            {
                                colContentLength[col] = DGVDTA.Rows[row].Cells[col].Value.ToString().Length;
                            }
                        }

                    }
                }

                for (int j = 0; j < DGVDTA.RowCount - 1; j++)//写内容
                {
                    string tempStr = "";
                    for (int k = 0; k < DGVDTA.ColumnCount; k++)
                    {
                        if (k > 0)
                        {
                            tempStr += "\t";
                        }

                        if (DGVDTA.Rows[j].Cells[k].Value == null)
                        {
                            tempStr += "".PadRight(colContentLength[k]);
                        }
                        else
                        {
                            tempStr += DGVDTA.Rows[j].Cells[k].Value.ToString().PadRight(colContentLength[k]);
                        }
                    }
                    sw.WriteLine(tempStr);
                }
                //fileSaved = true;
            }

            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }


            finally
            {
                sw.Close();

            }


        }
        public void FtpExportToTxt(DataGridView DGVTCJ, string Fn)//导出到本地硬盘
        {

            GarbageCollect();
            FlushMemory();
            string Filename;

            Filename = @Application.StartupPath + "\\serialTxt\\" + Fn.ToString();
            StreamWriter sw = new StreamWriter(Filename, false, System.Text.Encoding.GetEncoding("GB2312"));

            try
            {
                int[] colContentLength = new int[DGVTCJ.ColumnCount];//控制列宽
                for (int col = 0; col < DGVTCJ.ColumnCount; col++)
                {
                    colContentLength[col] = DGVTCJ.Columns[col].HeaderText.Length;
                }

                for (int row = 0; row < DGVTCJ.RowCount - 1; row++)
                {
                    for (int col = 0; col < DGVTCJ.ColumnCount; col++)
                    {
                        if (DGVTCJ.Rows[row].Cells[col].Value != null)
                        {
                            if (DGVTCJ.Rows[row].Cells[col].Value.ToString().Length > colContentLength[col])
                            {
                                colContentLength[col] = DGVTCJ.Rows[row].Cells[col].Value.ToString().Length;
                            }
                        }

                    }
                }

                for (int j = 0; j < DGVTCJ.RowCount - 1; j++)//写内容
                {
                    string tempStr = "";
                    for (int k = 0; k < DGVTCJ.ColumnCount; k++)
                    {
                        if (k > 0)
                        {
                            tempStr += "\t";
                        }

                        if (DGVTCJ.Rows[j].Cells[k].Value == null)
                        {
                            tempStr += "".PadRight(colContentLength[k]);
                        }
                        else
                        {
                            tempStr += DGVTCJ.Rows[j].Cells[k].Value.ToString().PadRight(colContentLength[k]);
                        }
                    }
                    sw.WriteLine(tempStr);
                }
                //fileSaved = true;
            }

            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }


            finally
            {
                sw.Close();

            }
        }
        #region DateGridView导出到csv格式的Excel
        /// <summary>  
        /// 常用方法，列之间加\t，一行一行输出，此文件其实是csv文件，不过默认可以当成Excel打开。  
        /// </summary>  
        /// <remarks>  
        /// using System.IO;  
        /// </remarks>  
        /// <param name="dgv"></param>  
        private void DGVToExcel(DataGridView DGVTCJ, string Fn)
        {

            GarbageCollect();
            FlushMemory();
            string Filename;

            Filename = @Application.StartupPath + "\\serialTxt\\" + Fn.ToString();
            StreamWriter sw = new StreamWriter(Filename, false, System.Text.Encoding.GetEncoding("GB2312"));

            try
            {
                int[] colContentLength = new int[DGVTCJ.ColumnCount];//控制列宽
                for (int col = 0; col < DGVTCJ.ColumnCount; col++)
                {
                    colContentLength[col] = DGVTCJ.Columns[col].HeaderText.Length;
                }

                for (int row = 0; row < DGVTCJ.RowCount - 1; row++)
                {
                    for (int col = 0; col < DGVTCJ.ColumnCount; col++)
                    {
                        if (DGVTCJ.Rows[row].Cells[col].Value != null)
                        {
                            if (DGVTCJ.Rows[row].Cells[col].Value.ToString().Length > colContentLength[col])
                            {
                                colContentLength[col] = DGVTCJ.Rows[row].Cells[col].Value.ToString().Length;
                            }
                        }

                    }
                }

                for (int j = 0; j < DGVTCJ.RowCount - 1; j++)//写内容
                {
                    string tempStr = "";
                    for (int k = 0; k < DGVTCJ.ColumnCount; k++)
                    {
                        if (k > 0)
                        {
                            tempStr += "\t";
                        }

                        if (DGVTCJ.Rows[j].Cells[k].Value == null)
                        {
                            tempStr += "".PadRight(colContentLength[k]);
                        }
                        else
                        {
                            tempStr += DGVTCJ.Rows[j].Cells[k].Value.ToString().PadRight(colContentLength[k]);
                        }
                    }
                    sw.WriteLine(tempStr);
                }
                //fileSaved = true;
            }

            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }


            finally
            {
                sw.Close();

            }
        }
        #endregion  
        //调用以下函数 
        private FtpStatusCode UploadFun(string fileName, string uploadUrl)
        {
            GarbageCollect();
            FlushMemory();
            Stream requestStream = null;
            FileStream fileStream = null;
            FtpWebResponse uploadResponse = null;
            try
            {
                FtpWebRequest uploadRequest =
                (FtpWebRequest)WebRequest.Create(uploadUrl);
                uploadRequest.Method = WebRequestMethods.Ftp.UploadFile;

                uploadRequest.Proxy = null;
                NetworkCredential nc = new NetworkCredential();

                string APP_Path = Application.StartupPath;//获取启动了应用程序的可执行文件的路径，不包括可执行文件的名称            
                string XmlFile = APP_Path + "\\SerialConfig.xml";


                string uid = Helper_Xml.Read(XmlFile, "/Root/DSI/UserInfo", "Uid").ToString();
                string pwd = Helper_Xml.Read(XmlFile, "/Root/DSI/UserInfo", "Pwd").ToString();
                nc.UserName = uid;
                nc.Password = pwd;

                uploadRequest.Credentials = nc; //修改getCredential();错误2


                requestStream = uploadRequest.GetRequestStream();
                fileStream = File.Open(fileName, FileMode.Open);

                byte[] buffer = new byte[1024];
                int bytesRead;
                while (true)
                {
                    bytesRead = fileStream.Read(buffer, 0, buffer.Length);
                    if (bytesRead == 0)
                        break;
                    requestStream.Write(buffer, 0, bytesRead);
                }
                requestStream.Close();

                uploadResponse = (FtpWebResponse)uploadRequest.GetResponse();
                return uploadResponse.StatusCode;

            }
            catch (UriFormatException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (WebException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (uploadResponse != null)
                    uploadResponse.Close();
                if (fileStream != null)
                    fileStream.Close();
                if (requestStream != null)
                    requestStream.Close();
            }
            return FtpStatusCode.Undefined;
        }
        //删除FTP上的文件


        [DllImport("kernel32.dll")]
        public static extern bool SetProcessWorkingSetSize(IntPtr process, int minSize, int maxSize);

        public static void GarbageCollect()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        public static void FlushMemory()
        {
            GarbageCollect();

            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {

                SetProcessWorkingSetSize(System.Diagnostics.Process.GetCurrentProcess().Handle, -1, -1);
            }
        }

        //发送电子邮件成功返回True，失败返回False
        private bool SendMail()
        {

            string APP_Path = Application.StartupPath;//获取启动了应用程序的可执行文件的路径，不包括可执行文件的名称            
            string XmlFile = APP_Path + "\\SerialConfig.xml";


            string mailto = Helper_Xml.Read(XmlFile, "/Root/Send/TO", "Mail").ToString();
            string mailtoname = Helper_Xml.Read(XmlFile, "/Root/Send/TO", "Name").ToString();
            string mailcc = Helper_Xml.Read(XmlFile, "/Root/Send/CC", "Mail").ToString();
            string mailccname = Helper_Xml.Read(XmlFile, "/Root/Send/CC", "Name").ToString();
            string mailfrom = Helper_Xml.Read(XmlFile, "/Root/Send/FR", "Mail").ToString();
            string mailfromname = Helper_Xml.Read(XmlFile, "/Root/Send/FR", "Name").ToString();
            string mailsubject = Helper_Xml.Read(XmlFile, "/Root/Send/SU", "Subject").ToString();
            //MailAddress from = new MailAddress("cjh@teac.com.cn", "电脑课程建红");
            //收件人地址
            MailAddress to = new MailAddress(mailto, mailtoname);
            MailAddress cc = new MailAddress(mailcc, mailccname);

            MailMessage message = new MailMessage();

            message.To.Add(to);
            //message.CC.Add(cc);
            message.From = new MailAddress(mailfrom, mailfromname, System.Text.Encoding.UTF8);


            //添加附件，判断文件存在就添加
            //if (System.IO.File.Exists(this.txtAttachment.Text))
            //{
            //    Attachment item = new Attachment(this.txtAttachment.Text, MediaTypeNames.Text.Plain);
            //    message.Attachments.Add(item);
            //}
            message.Subject = mailsubject + ',' + TotalTime; // 设置邮件的标题
            if (uploadlen != 0)
            {
                message.Body = "Dear All,\r\n" + uploadlen + "のレコードがテキストファイル（sndta.txt）に書き出されました。\r\n" + "ファイル（sndta.txt）がTCJのFTPサーバにアップロード完了されました。\r\n" + "ご確認お願い致します。\r\n「" + Helper_Hard.GetComputerName() + "\r\n" + Helper_Hard.GetIPAddress() + "\r\n" + Helper_Hard.GetUserName() + "\r\n" + DateTime.Now.ToString() + "」\r\n" + "このメールはシステムより自動送信されています。\r\n返信は受付できませんので、ご了承ください。\r\n\n\r\n\n";  //发送邮件的正文
            }
            else
            {
                message.Body = "Dear All,\r\n" + "本日出荷数量は0件ですので、テキストファイル（sndta.txt）に書き出されませんでした。\r\n" + "ファイル（sndta.txt）がTCJのFTPサーバにアップロードされませんでした。\r\n" + "ご確認お願い致します。\r\n「" + Helper_Hard.GetComputerName() + "\r\n" + Helper_Hard.GetIPAddress() + "\r\n" + Helper_Hard.GetUserName() + "\r\n" + DateTime.Now.ToString() + "」\r\n" + "このメールはシステムより自動送信されています。\r\n返信は受付できませんので、ご了承ください。\r\n\n";  //发送邮件的正文
            }
            message.BodyEncoding = System.Text.Encoding.Default;
            //MailAddress other = new MailAddress("davische@teac.com.cn");
            //message.CC.Add(other); //添加抄送人
            //创建一个SmtpClient 类的新实例,并初始化实例的SMTP 事务的服务器
            SmtpClient client = new SmtpClient(@"192.168.16.254");
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.EnableSsl = false;
            //身份认证
            client.Credentials = new System.Net.NetworkCredential("itsup@teac.com.cn", "k2ofc0bj");
            bool ret = true; //返回值
            try
            {
                client.Send(message);
            }
            catch (SmtpException ex)
            {
                MessageBox.Show(ex.Message);
                ret = false;
            }
            catch (Exception ex2)
            {
                MessageBox.Show(ex2.Message);
                ret = false;
            }
            return ret;
        }
        protected bool getTimeSpan(string timeStr)
        {
            //判断当前时间是否在工作时间段内
            string _strWorkingDayAM = "08:00";//工作时间上午08:30
            string _strWorkingDayPM = "13:30";
            TimeSpan dspWorkingDayAM = DateTime.Parse(_strWorkingDayAM).TimeOfDay;
            TimeSpan dspWorkingDayPM = DateTime.Parse(_strWorkingDayPM).TimeOfDay;

            //string time1 = "2017-2-17 8:10:00";
            DateTime t1 = Convert.ToDateTime(timeStr);

            TimeSpan dspNow = t1.TimeOfDay;
            if (dspNow > dspWorkingDayAM && dspNow < dspWorkingDayPM)
            {
                return true;
            }
            return false;
        }
    }
}


