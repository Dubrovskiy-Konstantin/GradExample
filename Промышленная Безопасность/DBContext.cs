using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Промышленная_Безопасность
{
    /// <summary>
    /// Контекст для работы с базой данных
    /// </summary>
    public class DBContext
    {
        private readonly string _connectionString;

        /// <summary>
        /// Контекст для работы с базой данных
        /// </summary>
        public DBContext()
        {
            XmlDocument xmlDcoument = new XmlDocument();
            xmlDcoument.Load(@"AppSettings.xml");
            this._connectionString = xmlDcoument.SelectSingleNode("Settings").SelectSingleNode("ConnectionString").InnerText;
        }

        /// <summary>
        /// Универсальный метод для работы с БД с помощью sql запросов
        /// </summary>
        /// <param name="operations">Запрос T-SQL (оидн или несколько)</param>
        /// <returns>Табилца объектов результата</returns>
        public List<List<object>> Execute(params string[] operations)
        {
            var result = new List<List<object>>();
            try
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    if (operations != null)
                    {
                        foreach (var operation in operations)
                        {
                            SqlCommand cmd = new SqlCommand(operation, connection);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var row = new List<object>();
                                    for (int i = 0; i < reader.FieldCount; i++)
                                    {
                                        row.Add(reader.GetValue(i));
                                    }
                                    result.Add(row);
                                }
                            }
                        }
                    }

                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return result;
        }

        /// <summary>
        /// Взять данные из таблицы как они хранятся
        /// </summary>
        /// <param name="sql">Запрос sql-таблицы</param>
        /// <returns>DataSet результата</returns>
        public DataSet GetDataSet(string sql)
        {
            DataSet dataSet = null;
            SqlDataAdapter adapter;

            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    adapter = new SqlDataAdapter(sql, connection);
                    dataSet = new DataSet();

                    adapter.Fill(dataSet);
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return dataSet;
        }
    }
}
