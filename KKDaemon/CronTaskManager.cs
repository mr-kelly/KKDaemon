using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using NCrontab;

namespace KKDaemon
{
    public delegate void CronDelegate();

    public class CronTask
    {
        public DateTime LastTime;
        public CronDelegate Callback;
    }

    class CronTaskManager
    {
        Dictionary<CrontabSchedule, CronTask> Crons = new Dictionary<CrontabSchedule, CronTask>();
        public CronTaskManager()
        {
            var thread = new Thread(() =>
            {
                while (true)
                {

                    foreach (var cron in Crons.Keys)
                    {
                        var task = Crons[cron];
                        var occurs = cron.GetNextOccurrences(task.LastTime, DateTime.Now);
                        foreach (var occur in occurs)
                        {
                            Crons[cron].LastTime = DateTime.Now;
                            Crons[cron].Callback();
                            continue;
                        }

                    }
                    Thread.Sleep(1000);
                }
            });
            thread.Start();
        }
        public void BeginTask(string cronDesc, CronDelegate callback)
        {
            Crons.Add(CrontabSchedule.Parse(cronDesc), new CronTask (){
                LastTime = DateTime.Now,
                Callback = callback,
            });
        }
    }

    //public class MyTask : StandardTask
    //{

    //    public override bool Execute(Dictionary<string, object> customData)
    //    {
    //        Console.WriteLine("Hello from MyTask");
    //        return true;
    //    }
    //}
}
