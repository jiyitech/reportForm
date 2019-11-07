using ServiceStack.DataAnnotations;
using ServiceStack.OrmLite;
using System.Linq;
using System.Reflection;
using System.Text;

namespace reportform
{
    public static class SqliteTableFactory
    {
        // 项目数据库全局
        static public string GobalPrefix { get; set; }

        static public string connectionString { get; set; }

        public static bool GenerateTable<T>()
        {
            var dbFactory = new OrmLiteConnectionFactory(connectionString, SqliteDialect.Provider);
            using (var db = dbFactory.Open())
            {
                // 获取实体对应数据库表名
                AliasAttribute attribute = (AliasAttribute)typeof(T).GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault();
                var tableName = attribute != null ? attribute.Name : typeof(T).Name;
                if (!db.CreateTableIfNotExists<T>())
                {
                    foreach (PropertyInfo prop in typeof(T).GetProperties())
                    {
                        //AliasAttribute attribute = (AliasAttribute)typeof(T).GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault();
                        if (!db.ColumnExists(prop.Name, attribute == null ? typeof(T).Name : attribute.Name)) //= false
                        {
                            var modelDef = ModelDefinition<T>.Definition;
                            if (modelDef.IgnoredFieldDefinitions.FindAll(x => x.FieldName == prop.Name).Count == 0)
                            {
                                var fieldDef = modelDef.GetFieldDefinition(prop.Name);
                                db.AddColumn(typeof(T), fieldDef);
                            }
                        }
                    }
                }
                //if (attribute.Name == "Log")
                //{ return true; }

                #region 全局前缀字段

                //// 执行写入项目全局前缀id
                //StringBuilder sql = new StringBuilder();
                //// 查询数据表中是否存在 该字段
                //sql.Append($" if exists(select * from syscolumns where id=object_id('{tableName}') and name='id') ");
                //// 删除  现有计算列
                //sql.Append($" alter table {tableName} drop column id ");
                //// 添加  计算列
                //sql.Append($@" 
                //              alter table {tableName} add id as '{GobalPrefix}'+CONVERT([nvarchar],numberId) PERSISTED
                //             ");

                //var ttt = sql.ToString();
                //// 执行
                //db.ExecuteSql(sql.ToString());

                #endregion

                return true;
            }
        }
    }
}
