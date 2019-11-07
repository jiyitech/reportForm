using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
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
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            //SqliteTableFactory.connectionString = "Data Source=E:\\reportforms\\DB\\hehe.db;version=3";
            //SqliteTableFactory.GenerateTable<ReportInfo>();//增加表
            try
            {
                UnhandledExceptionCatch.CatchUnhandledException();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
            catch (Exception ex)
            {

                throw ex;
            }
            
        }
    }
}
