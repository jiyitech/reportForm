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
        string autoname;
        string totalname;
        string address;
        SqliteConnect con;
        public Form1()
        {
            InitializeComponent();
            frm = this;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog savefile = new SaveFileDialog();
            //如果文件名未写后缀名则自动添加     *.*不会自动添加后缀名
            savefile.AddExtension = true;
            savefile.Filter = "|*.*";
            savefile.FileName = "斗轮机全自动投用率报表.xlsx";
            if (DialogResult.OK == savefile.ShowDialog())
            {
                Excel.WriteSheet(savefile.FileName, myBindingSource.DataSource as DataTable);
            }
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
            selectFromIFIX(autoname, totalname);
            showTable(address);
        }

        public void selectFromIFIX(string autoName, string totalName)
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
                string autoTime = JsonConvert.SerializeObject(values[0].Value);
                string totalTime = JsonConvert.SerializeObject(values[1].Value);
                insertTable(autoTime, totalTime);//插入数据库
            }
        }

        public void insertTable(string autoTime, string totalTime)
        {
            ReportInfo report = new ReportInfo();

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
            double.TryParse(autoTime, out double autoEnd);
            double.TryParse(totalTime, out double totalEnd);
            report.autoEnd = autoEnd;
            report.autoTime = autoEnd - previousAuto;
            report.totalEnd = totalEnd;
            report.totalTime = totalEnd - previousTotal;
            report.percent = (autoEnd - previousAuto) / (totalEnd - previousTotal) * 100;
            report.date = DateTime.Now;
            con.insert(report);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
            con = new SqliteConnect();
            string jsonfile = Application.StartupPath + "\\appsetting.json";
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
            showTable(address);
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

        private void button2_Click(object sender, EventArgs e)
        {
            var startTime = dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss");
            var endTime = dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss");
            var result = con.selectByTime(startTime, endTime);
            dataGridView1.BeginInvoke(new Action(() =>
            {
                try
                {
                    myBindingSource.DataSource = result.ToList();
                    dataGridView1.DataSource = myBindingSource.DataSource;//将BindingSource绑定到GridView
                    dataGridView1.Columns["date"].HeaderText = "录入日期";
                    dataGridView1.Columns["dutyName"].HeaderText = "班值";
                    dataGridView1.Columns["autoTime"].HeaderText = "自动运行时间";
                    dataGridView1.Columns["autoEnd"].HeaderText = "上次自动运行结束时间";
                    dataGridView1.Columns["totalTime"].HeaderText = "运行总时间";
                    dataGridView1.Columns["totalEnd"].HeaderText = "上次运行结束总时间";
                    dataGridView1.Columns["percent"].HeaderText = "%";
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }));
        }
    }
}
