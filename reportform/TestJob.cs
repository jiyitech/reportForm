
using Quartz;
using Quartz.Impl;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Threading.Tasks;

namespace reportform
{
    public class TestJob : IJob
    {


        public MyDelegateHandler myDelegate = null;

        /// <summary>
        /// 作业调度定时执行的方法
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public async Task Execute(IJobExecutionContext context)
        {
            try
            {
                Form1.frm.test();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
        }
    }
}
