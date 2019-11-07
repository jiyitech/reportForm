using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Castle.Core.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using TitaniumAS.Opc.Client.Common;
using TitaniumAS.Opc.Client.Da;

namespace reportform
{
    public delegate void MyDelegateHandler();
    public partial class Form1 : Form
    {
        SQLiteDataAdapter mAdapter;
        DataTable mTable;
        public static Form1 frm;
        BindingSource myBindingSource = new BindingSource();//创建BindingSource
        public Form1()
        {
            InitializeComponent();
            frm = this;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string a = "我是文件.xls";
            ExportExcels(a, dataGridView1);
        }

        /// </summary>
        /// <param name="fileName">文件路径</param>
        /// <param name="myDGV">控件DataGridView</param>
        private void ExportExcels(string fileName, DataGridView myDGV)
        {
            string saveFileName = "";
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xls";
            saveDialog.Filter = "Excel文件|*.xls";
            saveDialog.FileName = fileName;
            saveDialog.ShowDialog();
            saveFileName = saveDialog.FileName;
            if (saveFileName.IndexOf(":") < 0) return; //被点了取消
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("无法创建Excel对象，可能您的机子未安装Excel");
                return;
            }
            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1
                                                                                                                                  //写入标题
            for (int i = 0; i < myDGV.ColumnCount; i++)
            {
                worksheet.Cells[1, i + 1] = myDGV.Columns[i].HeaderText;
            }
            //写入数值
            for (int r = 0; r < myDGV.Rows.Count; r++)
            {
                for (int i = 0; i < myDGV.ColumnCount; i++)
                {
                    worksheet.Cells[r + 2, i + 1] = myDGV.Rows[r].Cells[i].Value;
                }
                System.Windows.Forms.Application.DoEvents();
            }
            worksheet.Columns.EntireColumn.AutoFit();//列宽自适应
            if (saveFileName != "")
            {
                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(saveFileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                }
            }
            xlApp.Quit();
            GC.Collect();//强行销毁
            MessageBox.Show("文件： " + fileName + ".xls 保存成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void showTable(string address)
        {
            mAdapter = new SQLiteDataAdapter("select * from ReportInfo", new SQLiteConnection($"Data Source={address};version=3"));
            mTable = new DataTable(); // Don't forget initialize!
            mAdapter.Fill(mTable);
            // 绑定数据到DataGridView
            dataGridView1.BeginInvoke(new Action(() =>
            {
                try
                {
                    myBindingSource.DataSource = mTable;
                    dataGridView1.DataSource = myBindingSource.DataSource;//将BindingSource绑定到GridView
                    dataGridView1.Columns["date"].HeaderText = "录入日期";
                    dataGridView1.Columns["dutyName"].HeaderText = "班值";
                    dataGridView1.Columns["autoTime"].HeaderText = "自动运行时间";
                    dataGridView1.Columns["autoEnd"].HeaderText = "上次自动运行结束时间";
                    dataGridView1.Columns["totalTime"].HeaderText = "运行总时间";
                    dataGridView1.Columns["totalEnd"].HeaderText = "上次运行结束总时间";
                    dataGridView1.Columns["percent"].HeaderText = "%";
                    DataGridViewColumn col = dataGridView1.Columns[0];
                    // 按降序(即始终每次新添加的数据排最前)
                    ListSortDirection direction = ListSortDirection.Descending;
                    dataGridView1.Sort(col, direction); // 执行指定排序规则
                }
                catch (Exception e)
                {
                    throw e;
                }
            }));
        }

        public void test()
        {
            string jsonfile = Application.StartupPath + "\\appsetting.json";
            string autoname;
            string totalname;
            string address;
            using (System.IO.StreamReader file = System.IO.File.OpenText(jsonfile))
            {
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    JObject o = (JObject)JToken.ReadFrom(reader);
                    autoname = o["inputdata"]["autoName"].ToString();
                    totalname = o["inputdata"]["totalName"].ToString();
                    address = o["inputdata"]["address"].ToString();
                }
            }
            selectFromIFIX(autoname, totalname);
            showTable(address);
        }

        public void selectFromIFIX(string autoName,string totalName)
        {
            Uri url = UrlBuilder.Build("Intellution.OPCiFIX.1");
            using (var server = new OpcDaServer(url))
            {
                // Connect to the server first.
                server.Connect();
                OpcDaGroup group = server.AddGroup("MyGroup");
                group.IsActive = true;

                var definition1 = new OpcDaItemDefinition
                {
                    ItemId = autoName,
                    IsActive = true
                };
                var definition2 = new OpcDaItemDefinition
                {
                    ItemId = totalName,
                    IsActive = true
                };
                OpcDaItemDefinition[] definitions = { definition1, definition2 };
                OpcDaItemResult[] results = group.AddItems(definitions);

                // Handle adding results.
                foreach (OpcDaItemResult result in results)
                {
                    if (result.Error.Failed)
                        MessageBox.Show($"Error adding items: {result.Error}");
                }
                OpcDaItemValue[] values = group.Read(group.Items, OpcDaDataSource.Device);
                string value1 = JsonConvert.SerializeObject(values[0].Value);
                string value2 = JsonConvert.SerializeObject(values[1].Value);
                insertTable(value1, value2);//插入数据库
            }
        }

        public void insertTable(string value1,string value2)
        {
            ReportInfo report = new ReportInfo();
            SqliteConnect con = new SqliteConnect();
            IEnumerable<ReportInfo> resNew = con.selectNew();
            double? previousAuto;
            double? previousTotal;
            if (resNew.Count() == 0)
            {
                previousAuto = 0;
                previousTotal = 0;
            }
            else
            {
                previousAuto = resNew.ToList()[0].autoEnd;
                previousTotal = resNew.ToList()[0].totalEnd;
            }
            int currentTime = DateTime.Now.Hour;
            if (currentTime > 0 && currentTime <= 8)
            {
                report.dutyName = "日";
            }
            else if (currentTime > 8 && currentTime <= 16)
            {
                report.dutyName = "中";
            }
            else if (currentTime > 16)
            {
                report.dutyName = "夜";
            }
            report.autoEnd = double.Parse(value1);
            report.autoTime = double.Parse(value1) - previousAuto;
            report.totalEnd = double.Parse(value2);
            report.totalTime = double.Parse(value2) - previousTotal;
            report.percent = (double.Parse(value1) - previousAuto) / (double.Parse(value2) - previousTotal) * 100;
            con.insert(report);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
            TaskInit.Init();
        }


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //取消关闭窗口
            e.Cancel = true;
            //最小化主窗口
            this.WindowState = FormWindowState.Minimized;
            //不在系统任务栏显示主窗口图标
            this.ShowInTaskbar = false;
            //提示气泡
            notifyIcon1.ShowBalloonTip(2000, "最小化到托盘", "程序已经缩小到托盘，单击打开程序。", ToolTipIcon.Info);
        }


        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (this.WindowState == FormWindowState.Minimized)
                {
                    //还原窗体
                    this.WindowState = FormWindowState.Normal;
                    //系统任务栏显示图标
                    this.ShowInTaskbar = true;
                }
                //激活窗体并获取焦点
                this.Activate();
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Dispose();
            this.Close();
        }

    }
}
