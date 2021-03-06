﻿
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
    public class ExcelJob : IJob
    {
        /// <summary>
        /// 作业调度定时执行的方法
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public async Task Execute(IJobExecutionContext context)
        {
            try
            {
                string time = context.NextFireTimeUtc.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");
                Form1.frm.autoGenerateExcel(); 
                Form1.frm.setMonthTime(time);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
        }
    }
}
