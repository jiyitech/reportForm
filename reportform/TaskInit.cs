using Quartz;
using Quartz.Impl;
using Quartz.Impl.Calendar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace reportform
{


    public static class TaskInit
    {
        public static void Init()
        {
            ScheduleInit().GetAwaiter();
        }
        public static async Task ScheduleInit()
        {

            // 关于quartznet的更多文档请看https://www.quartz-scheduler.net/documentation/quartz-3.x/tutorial/index.html
            // 创建作业调度器
            ISchedulerFactory factory = new StdSchedulerFactory();
            IScheduler scheduler = await factory.GetScheduler();
            // 开始运行调度器
            await scheduler.Start();
            // 创建作业
            IJobDetail testJob = JobBuilder.Create<TestJob>()
                .Build();
            // 创建定时任务
            ITrigger jobTrigger = TriggerBuilder.Create()
                .WithIdentity("job1", "triggerGroup")
                .StartNow()
                .WithSimpleSchedule(x => x
                .WithIntervalInSeconds(10)
                .RepeatForever())
                .Build();

            // 创建调度任务
            ITrigger cronTrigger = TriggerBuilder.Create()
                .WithIdentity("job2", "triggerGroup")
                .StartNow()
                .WithCronSchedule("0 0 0,8,16 * * ?")
                .ForJob(testJob)
                .Build();
            // 任务加入调度器
            await scheduler.ScheduleJob(testJob, jobTrigger);
            await scheduler.ScheduleJob(cronTrigger);
        }
    }
}
