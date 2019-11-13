using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Dapper;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ServiceStack.OrmLite;
using ServiceStack.OrmLite.Sqlite;

namespace reportform
{

     class SqliteConnect
    {
        private IDbConnection db;
        public SqliteConnect()
        {
            db = new SQLiteConnection(LoadConnectString());
            db.Open();
        }

        private string LoadConnectString()
        {
            string jsonfile = Application.StartupPath + "\\appsetting.json";
            using (System.IO.StreamReader file = System.IO.File.OpenText(jsonfile))
            {
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    JObject o = (JObject)JToken.ReadFrom(reader);
                    string address = o["inputdata"]["address"].ToString();
                    string compose = $"Data Source = { address }; version = 3";
                    return compose;
                }
            }
        }

        public IEnumerable<ReportInfo> selectNew(string bucketWheelName)
        {
           return db.Query<ReportInfo>($"select * from ReportInfo where id = (select MAX(Id) from ReportInfo where bucketWheel = '{bucketWheelName}')");
        }

        public IEnumerable<ReportInfo> selectByTimeAndName(string startTime, string endTime, string bucketWheelName)
        {
            return db.Query<ReportInfo>($"select * from ReportInfo where date >='{startTime}' and date<='{endTime}' and bucketWheel='{bucketWheelName}'order by date desc");
        }

        public IEnumerable<ReportInfo> selectByTime(string startTime, string endTime)
        {
            return db.Query<ReportInfo>($"select * from ReportInfo where date >='{startTime}' and date<='{endTime}' order by date desc");
        }

        public long insert(ReportInfo report)
        {
            OrmLiteConfig.DialectProvider = SqliteOrmLiteDialectProvider.Instance;//OrmLiteConfig.DialectProvider 静态属性，我们使用前必须赋予初始值
            return db.Insert<ReportInfo>(report);
        }
    }
}
