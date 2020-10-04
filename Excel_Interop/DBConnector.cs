using System;
using System.Collections.Generic;
using System.Text;
using MySql.Data.MySqlClient;
using System.Reflection;
using refl = System.Reflection;
using System.Linq;
namespace Excel_Interop
{
    class DBConnector
    {
        delegate T Executor<T>(string str);
        delegate object? Invoker(object? instance, object?[]? parameters);

        public DBConnector(
            string host = "192.168.1.39",
            string database = "eds_db",
            string user = "root",
            string password = "3rfbR6RBM"
            )
        {
            Host = host;
            Dbase = database;
            User = user;
            Password = password;
            DBConnect();
        }

        string Dbase { get; set; }
        string Host { get; set; }
        string Password { get; set; }
        string User { get; set; }
        string CnnStr => $"Database={Dbase};Datasource={Host};User={User};Password={Password};";

        MySqlConnection Connection { get; set; }
        public void DBConnect()
        {
            Connection ??= new MySqlConnection(CnnStr);
            if (Connection.State == System.Data.ConnectionState.Closed) Connection.Open();
        }
        

        public IEnumerable<T> SynchronizeTable<T>(string tableName = "")
        {
            var type = typeof(T);
            var instance = (T)Activator.CreateInstance(type);
            var listOfInstance = new List<T>();
            var propertyName = "";
            var columnName = "";
            var typeName = type.Name;
            var value = new object();
            var getKeyValuePairs = new Invoker(type.GetMethod("get_KeyValuePairs").Invoke);
            var dict = (Dictionary<string, string>)getKeyValuePairs.Invoke(instance, new object[] { });
            var query = new MySqlCommand($"SELECT * FROM {Dbase}.{(tableName == "" ? typeName + "s" : tableName)}", Connection);
            var reader = query.ExecuteReader();
            Invoker invokeClassMethods;
            
            while (reader.Read())
            {
                instance = (T)Activator.CreateInstance(type);
                for (int i = 0; i < reader.GetSchemaTable().Rows.Count - 1; i++)
                {
                    columnName = (string)reader.GetSchemaTable().Rows[i]["ColumnName"];
                    if (!dict.TryGetValue(columnName, out propertyName)) continue;

                    invokeClassMethods = new Invoker(type.GetMethod($"set_{propertyName}").Invoke);
                    value = reader.GetValue(i);
                    if (value is DBNull) continue; 
                    invokeClassMethods(instance, new object[]{ value });
                    
                }
                listOfInstance.Add(instance);
            }
            reader.Close();
            return listOfInstance;



            //foreach (var method in methods)
            //{
            //    Console.WriteLine(method.Name);
            //}
        }

        ~DBConnector()
        {
            Connection.Close();
        }
    }
}
