using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MNV.ExcelDataToDB.DbContext
{
    public class MariaDbContext : IDisposable
    {
        private string _connectionString;
        MySqlConnection _mySqlConnection;
        MySqlCommand _mySqlCommand;
        public MariaDbContext(string connectionString)
        {
            _connectionString = connectionString;
            _mySqlConnection = new MySqlConnection(_connectionString);
            _mySqlCommand = _mySqlConnection.CreateCommand();
            _mySqlConnection.Open();
        }

        /// <summary>
        /// ExecuteReader
        /// </summary>
        /// <param name="commandText"></param>
        /// <returns></returns>
        /// Created By : NVMANH (20/10/2019)
        public MySqlDataReader ExecuteReader(string commandText)
        {
            _mySqlCommand.CommandText = commandText;
            return _mySqlCommand.ExecuteReader();
        }

        public void CloseConnection()
        {
            _mySqlConnection.Close();
        }

        /// <summary>
        /// Thêm param + value vào command text
        /// (Lưu ý tên param cần có tiền tố @)
        /// </summary>
        /// <param name="paramName"></param>
        /// <param name="value"></param>
        public void AddParmWithValue(string paramName, object value)
        {
            _mySqlCommand.Parameters.AddWithValue(paramName, value);
        }

        public bool HasParam(string paramName)
        {
            try
            {
                var param = _mySqlCommand.Parameters[paramName];
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Set value cho Param
        /// (Lưu ý tên param cần có tiền tố @)
        /// </summary>
        /// <param name="paramName"></param>
        /// <param name="value"></param>
        public void SetParamValue(string paramName, object value)
        {
            try
            {
                _mySqlCommand.Parameters[paramName].Value = value;
            }
            catch (Exception)
            {
                _mySqlCommand.Parameters.AddWithValue(paramName, value);
            }
        }
        /// <summary>
        /// ExecuteNonQuery
        /// </summary>
        /// <param name="commandText"></param>
        /// <returns></returns>
        /// Created By : NVMANH (20/10/2019)
        public int ExecuteNonQuery(string commandText)
        {
            if (_mySqlConnection.State == System.Data.ConnectionState.Closed)
                _mySqlConnection.Open();
            _mySqlCommand.CommandText = commandText;
            var result = _mySqlCommand.ExecuteNonQuery();
            return result;
        }
        public void Dispose()
        {
            _mySqlConnection.Close();
        }
    }
}
