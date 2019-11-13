using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using TitaniumAS.Opc.Client.Common;
using TitaniumAS.Opc.Client.Da;

namespace reportform
{
    public static class UnhandledExceptionCatch
    {
        public static void CatchUnhandledException()
        {
            try
            {
                //处理未捕获的异常
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                //处理UI线程异常
                Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
                //处理非UI线程异常
                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
            }
            catch (Exception ex)
            {
                //Log.Error(ex.Message);
            }
        }

        static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            //Log.Error(e.Exception.Message);
        }

        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            // Log.Error((e.ExceptionObject as Exception).Message);
        }
    }

    static class Program
    {
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll ", SetLastError = true)]
        static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
        [DllImport("user32.dll", EntryPoint = "ShowWindow", CharSet = CharSet.Auto)]
        public static extern int ShowWindow(IntPtr hwnd, int nCmdShow);
        public const int SW_RESTORE = 9;
        public static IntPtr formhwnd;
        static Form1 form = null;
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            string address;
            string jsonfile = Application.StartupPath + "\\appsetting.json";
            using (System.IO.StreamReader file = System.IO.File.OpenText(jsonfile))
            {
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    JObject o = (JObject)JToken.ReadFrom(reader);
                    address = o["inputdata"]["address"].ToString();
                }
            }
                SqliteTableFactory.connectionString = $"Data Source={address};version=3";
                SqliteTableFactory.GenerateTable<ReportInfo>();//增加表
            try
            {
                UnhandledExceptionCatch.CatchUnhandledException();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                string proc = Process.GetCurrentProcess().ProcessName;
                Process[] processes = Process.GetProcessesByName(proc);
                if (processes.Length <= 1)
                {
                    form = new Form1();
                    Application.Run(form);
                }
                else
                {
                    for (int i = 0; i < processes.Length; i++)
                    {
                        if (processes[i].Id != Process.GetCurrentProcess().Id)
                        {
                            if (processes[i].MainWindowHandle.ToInt32() == 0)
                            {
                                formhwnd = FindWindow(null, "斗轮机全自动投用率报表");
                                ShowWindow(formhwnd, SW_RESTORE);
                                SwitchToThisWindow(formhwnd, true);
                            }
                            else
                            {
                                SwitchToThisWindow(processes[i].MainWindowHandle, true);
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
