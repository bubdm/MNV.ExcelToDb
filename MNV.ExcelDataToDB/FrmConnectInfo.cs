using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using MNV.ExcelDataToDB.DbContext;
using System.Globalization;
using MNV.ExcelDataToDB.Entity;

namespace MNV.ExcelDataToDB
{
    public partial class FrmConnectInfo : Form
    {
        BackgroundWorker _worker;
        Workbook _xlWorkbook;
        List<Row> rowsErro = new List<Row>();
        static string _connectionString = string.Empty;
        int _excelRowCount = 0;
        int _progessCount = 0;
        int _numberImportErro = 0;
        int _numberImportSuccess = 0;
        string _processInfo = string.Empty;
        string _filePath = string.Empty;
        List<MariaDbColumn> _mariaDbColumns = new List<MariaDbColumn>();
        public FrmConnectInfo()
        {
            InitializeComponent();
            lblProgessInfo.Text = string.Empty;
            lblImportCount.Text = string.Empty;
            cbxDbType.SelectedIndex = 1;
            progressBar.Maximum = 100;
            progressBar.Step = 1;
            progressBar.Value = 0;
            lblTimeProgessInfo.Text = string.Empty;
        }

        private void frmConnectInfo_Load(object sender, EventArgs e)
        {
            ResetInfoProgess();
        }

        /// <summary>
        /// Connect to server and get list data
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                cbxDatabases.Items.Clear();
                var server = txtServerInfo.Text.Trim();
                var port = txtPort.Text.Trim();
                var userName = txtUserName.Text.Trim();
                var password = txtPassword.Text.Trim();
                var connectionServerString = String.Format("server={0};port={1};user={2};password={3};charset=utf8;", server, port, userName, password);
                using (MariaDbContext mariaDbContext = new MariaDbContext(connectionServerString))
                {
                    MySqlDataReader _sqlDataReader = mariaDbContext.ExecuteReader("SHOW DATABASES;");
                    while (_sqlDataReader.Read())
                    {
                        string dbName = "";
                        for (int i = 0; i < _sqlDataReader.FieldCount; i++)
                        {
                            dbName += _sqlDataReader.GetValue(i).ToString();
                            if (dbName != "information_schema" && dbName != "mysql" && dbName != "performance_schema" && dbName != "sys")
                            {
                                cbxDatabases.Items.Add(dbName);
                            }
                        }
                    }
                    MessageBox.Show("Kết nối tới máy chủ thành công!");
                    cbxDatabases.Enabled = true;
                }

            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Có lỗi xảy ra khi kết nối tới server MySql: \r \n" + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra khi kết nối tới server: \r \n" + ex.Message);
            }

        }

        /// <summary>
        /// select database and get list tables
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbxDatabases_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                cbxTableList.Items.Clear();
                var dbName = cbxDatabases.Text;
                var server = txtServerInfo.Text.Trim();
                var port = txtPort.Text.Trim();
                var userName = txtUserName.Text.Trim();
                var password = txtPassword.Text.Trim();
                _connectionString = String.Format("server={0};port={1};database={4};user={2};password={3};charset=utf8;", server, port, userName, password, dbName);
                MySqlConnection connection = new MySqlConnection(_connectionString);
                MySqlCommand command = connection.CreateCommand();
                command.CommandText = "SHOW TABLES;";
                MySqlDataReader Reader;
                connection.Open();
                Reader = command.ExecuteReader();
                while (Reader.Read())
                {
                    string tableName = "";
                    for (int i = 0; i < Reader.FieldCount; i++)
                        tableName = Reader.GetValue(i).ToString();
                    cbxTableList.Items.Add(tableName);
                }
                connection.Close();
                cbxTableList.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cbxTableList_SelectedIndexChanged(object sender, EventArgs e)
        {
            _mariaDbColumns.Clear();
            var dbName = cbxDatabases.Text;
            var tableName = cbxTableList.Text;
            using (MariaDbContext mariaDbContext = new MariaDbContext(_connectionString))
            {
                var commandText = String.Format(@"select ordinal_position as ColumnID,
                                                        column_name as ColumnName,
                                                        data_type as DataType,
                                                        case when numeric_precision is not null
                                                                    then numeric_precision
                                                            else character_maximum_length end as MaxLength,
                                                        case when datetime_precision is not null
                                                                    then datetime_precision
                                                            when numeric_scale is not null
                                                                    then numeric_scale
                                                            else 0 end as DataPrecision,
                                                        case when is_nullable = 'NO' then False else True end as IsNullable,
                                                        column_default as ColumnDefault
                                                    from information_schema.columns
                                                    where table_name = '{0}' -- put table name here
                                                    and table_schema = '{1}' -- put schema name here
                                                    order by ordinal_position;", tableName, dbName);
                MySqlDataReader sqlDataReader = mariaDbContext.ExecuteReader(commandText);
                while (sqlDataReader.Read())
                {
                    var mariaDbColumn = new MariaDbColumn();
                    string columnName = "";
                    for (int i = 0; i < sqlDataReader.FieldCount; i++)
                    {
                        columnName = sqlDataReader.GetValue(i).ToString();
                        var fieldName = sqlDataReader.GetName(i);
                        var fieldValue = sqlDataReader.GetValue(i);
                        var property = mariaDbColumn.GetType().GetProperty(fieldName);
                        if (fieldValue != System.DBNull.Value && property != null)
                        {
                            if (property.PropertyType == typeof(Guid))
                            {
                                property.SetValue(mariaDbColumn, Guid.Parse((string)fieldValue));
                            }
                            else if (property.PropertyType == typeof(bool))
                            {
                                property.SetValue(mariaDbColumn, Convert.ToBoolean(fieldValue));
                            }
                            else
                            {
                                property.SetValue(mariaDbColumn, fieldValue);
                            }
                        }
                    }
                    _mariaDbColumns.Add(mariaDbColumn);
                }
                lblTotalColumn.Text = _mariaDbColumns.Count.ToString();
                SetDataMappingInGrid();
            }
        }

        /// <summary>
        /// Choose file for import
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChooseFile_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                // Filter by Excel Worksheets
                dialog.Title = "Chọn tệp Excel muốn nhập khẩu dữ liệu";
                dialog.Filter = "Excel Worksheets|*.xls;*.xlsx";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    EnabledAllControlInForm(false);
                    // Reset grid:
                    dgvExcelData.Columns.Clear();
                    dgvExcelData.Rows.Clear();
                    ExcelColumn.excelColumns.Clear();

                    var filePath = string.Empty;
                    filePath = dialog.FileName;
                    _filePath = dialog.FileName;
                    txtFilePath.Text = filePath;
                    // Worker:

                    backgroundWorker.RunWorkerAsync();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Set map column and enabled import button
        /// </summary>
        private void SetColumnMapAndEnabledImport()
        {
            SetDataMappingInGrid();
            if (ExcelColumn.excelColumns.Count > 0 && _mariaDbColumns.Count > 0)
            {
                btnImport.Invoke(new System.Action(() => btnImport.Enabled = true));
            }
            else
            {
                //btnImport.Enabled = false;
                btnImport.Invoke(new System.Action(() => btnImport.Enabled = false));
            }
        }

        /// <summary>
        /// Set dữ liệu hiển thị trên Grid thực hiện mapping các cột giữa Excel và table trong database
        /// </summary>
        private void SetDataMappingInGrid()
        {
            try
            {
                var columnInDBCount = _mariaDbColumns.Count;
                var columnHeaderExcelCount = ExcelColumn.excelColumns.Count;

                // Map cột lần lượt theo vị trí trên
                for (int i = 0; i < columnInDBCount; i++)
                {
                    var colValueExcel = string.Empty;
                    if (i < columnHeaderExcelCount)
                    {
                        colValueExcel = ExcelColumn.excelColumns[i].Name;
                        _mariaDbColumns[i].ExcelColHeaderName = colValueExcel;
                        _mariaDbColumns[i].ExcelColPossition = i + 1;
                    }

                }
                BindingSource myDataListBindingSource = new BindingSource();
                myDataListBindingSource.DataSource = _mariaDbColumns;
                dgvColumnMapping.Invoke(new System.Action(() => dgvColumnMapping.DataSource = myDataListBindingSource));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro: " + ex.Message);
            }

        }

        /// <summary>
        /// Thực hiện đọc dữ liệu từ Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_DoWorkReadData(object sender, DoWorkEventArgs e)
        {
            try
            {

                var watch = new System.Diagnostics.Stopwatch();
                watch.Start();

                _worker = sender as BackgroundWorker;
                // Vị trí dòng xác định là tiêu đề của bảng dữ liệu:
                var rowHeaderPossition = (int)numColumnHeader.Value;
                _processInfo = "xử lý các cột";
                _progessCount = 0;

                //Encrypt the selected file. I'll do this later. :)
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                _xlWorkbook = xlApp.Workbooks.Open(_filePath);
                _Worksheet xlWorksheet = _xlWorkbook.Sheets[1];
                Range _xlRangeHeader = xlWorksheet.UsedRange;
                int colCount = _xlRangeHeader.Columns.Count;

                // Trường hợp chọn cột tiêu đề chủ động thì thực hiện gán lại vùng dữ liệu chứa tiêu đề:
                if (rowHeaderPossition > 0)
                    _xlRangeHeader = xlWorksheet.Range[xlWorksheet.Cells[rowHeaderPossition, 1], xlWorksheet.Cells[rowHeaderPossition, colCount]];


                // Build danh sách các cột dữ liệu có trong bảng Excel: 
                var rowForErroList = new Row(1);
                rowsErro.Add(rowForErroList);
                for (int j = 1; j <= colCount; j++)
                {
                    //write the value to the Grid  
                    if (_xlRangeHeader.Cells[1, j] != null && _xlRangeHeader.Cells[1, j].Value2 != null)
                    {
                        var value = _xlRangeHeader.Cells[1, j].Value2.ToString();
                        rowForErroList.Cells.Add(new Cell(1, j, value));
                        // Lấy dòng bắt đầu có dữ liệu đầu tiên làm tiêu đề nếu không nhập dòng Excel làm tiêu đề:
                        rowHeaderPossition = 1;
                        // Tạo các cột dữ liệu:
                        DataGridViewTextBoxColumn dgvTextBoxColumn = new DataGridViewTextBoxColumn();
                        dgvTextBoxColumn.HeaderText = String.Format("[{0}] \n {1}", (j), value);
                        dgvTextBoxColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        dgvExcelData.Invoke(new System.Action(() => dgvExcelData.Columns.Add(dgvTextBoxColumn)));
                        ExcelColumn.excelColumns.Add(new ExcelColumn(value, j));
                    }
                    // Theo dõi tiến trình:
                    _progessCount++;
                    if (_worker.CancellationPending == true)
                    {
                        e.Cancel = true;
                        return;
                    }

                    var totalTimes = watch.Elapsed.TotalSeconds;
                    var totalMinutes = (int)(totalTimes / 60);
                    var second = (int)(totalTimes % 60);
                    lblTimeProgessInfo.Invoke(new System.Action(() => lblTimeProgessInfo.Text = String.Format("Time progess: {0}(m) {1}(s)", totalMinutes, second)));
                    _worker.ReportProgress((_progessCount * 100) / colCount);
                }

                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background  
                Marshal.ReleaseComObject(_xlRangeHeader);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release  
                _xlWorkbook.Close();
                Marshal.ReleaseComObject(_xlWorkbook);

                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                SetDataMappingInGrid();
                watch.Stop();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #region IMPORT
        private void btnImport_Click(object sender, EventArgs e)
        {
            var hasExcelData = (dgvColumnMapping.Rows.Count > 0);
            var hasChooseDataTableInDB = (dgvExcelData.Rows.Count > 0);
            var msg = string.Empty;
            if (!hasExcelData && !hasChooseDataTableInDB)
                msg = @"Vui lòng kết nối tới CSDL và chọn bảng dữ liệu cần nhập khẩu. Sau đó chọn file exel có dữ liệu để và thực hiện nhập khẩu.";


            if (dgvColumnMapping.Rows.Count == 0 && string.IsNullOrEmpty(msg))
                msg = @"Vui lòng kết nối tới Database của bạn và chọn bảng dữ liệu cần nhập khẩu.";

            if (dgvExcelData.Rows.Count == 0 && string.IsNullOrEmpty(msg))
                msg = @"Vui lòng chọn File Excel muốn nhập khẩu và đảm bảo File có dữ liệu.";

            if (!string.IsNullOrEmpty(msg))
            {
                MessageBox.Show(msg);
                return;
            }
            lblProgessInfo.Text = "Bắt đầu nhập khẩu...";
            EnabledAllControlInForm(false);
            ResetInfoProgess();
            backgroundWorkerImport.RunWorkerAsync();
        }

        /// <summary>
        /// Thực hiện nhập khẩu dữ liệu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundWorker_DoImport(object sender, DoWorkEventArgs e)
        {
            try
            {
                // Bắt đầu tính thời gian:
                var watch = new System.Diagnostics.Stopwatch();
                watch.Start();

                // Khởi background thực hiện:
                _worker = sender as BackgroundWorker;
                _progessCount = 0;
                _processInfo = "nhập khẩu";
                _numberImportErro = 0;
                _numberImportSuccess = 0;


                // Đọc tệp nhập khẩu:
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                _xlWorkbook = xlApp.Workbooks.Open(_filePath);
                _Worksheet xlWorksheet = _xlWorkbook.Sheets[1];
                Range _xlRange = xlWorksheet.UsedRange;
                _excelRowCount = _xlRange.Rows.Count;
                int? _excelColCount = _xlRange.Columns.Count;
                var columnHeaderPossition = numColumnHeader.Value; // Cột xác định là tiêu đề của bảng dữ liệu:
                var totalRecord = _excelRowCount - 1; // Tổng số bản ghi trên tệp (trừ 1 bản ghi làm tiêu đề)

                // Khai báo các biến cần sử dụng trong quá trình nhập khẩu:
                var totalProgessCount = totalRecord + 4;
                var deleteOldData = cboRemoveOldData.Checked; // Xóa dữ liệu cũ

                var tableName = string.Empty; // Tên bảng dữ liệu sẽ import
                cbxTableList.Invoke(new System.Action(() => tableName = cbxTableList.Text));
                txtQuery.Invoke(new System.Action(() => txtQuery.Text = string.Empty));

                // Duyệt từng dòng dữ liệu, thực hiện sau đó import từng dòng vào database:
                for (int i = 2; i <= _excelRowCount; i++)
                {

                    var colInsertNames = string.Empty;
                    var colInsertValues = string.Empty;
                    var fields = new List<Field>(); // Chứa danh sách các giá trị sẽ insert vào database;

                    // Bắt đầu progess tiến trình import dữ liệu vào database:
                    using (MariaDbContext mariaDbContext = new MariaDbContext(_connectionString))
                    {
                        foreach (var col in _mariaDbColumns)
                        {
                            var fieldName = col.ColumnName;
                            var dataType = col.DataType;
                            var colExcelName = col.ExcelColHeaderName;
                            var colExcelPossition = col.ExcelColPossition;
                            var valueToImport = (object)string.Empty;
                            if (colExcelPossition != null && colExcelPossition <= _excelColCount)
                            {
                                valueToImport = _xlRange.Cells[i, colExcelPossition].Value2;
                            }

                            if (!string.IsNullOrWhiteSpace(colExcelName))
                            {
                                switch (dataType)
                                {
                                    case "int":
                                        int newValueInteger;
                                        if (valueToImport == null)
                                            newValueInteger = 0;
                                        else
                                            valueToImport = int.TryParse(valueToImport.ToString(), out newValueInteger) ? newValueInteger : 0;
                                        break;
                                    case "bit":
                                        int newValueBit;
                                        if (valueToImport == null)
                                            newValueBit = 0;
                                        else
                                            valueToImport = int.TryParse(valueToImport.ToString(), out newValueBit) ? newValueBit : 0;
                                        break;
                                    case "date":
                                        DateTime date;
                                        if (valueToImport == null)
                                        {
                                            valueToImport = DateTime.Now.ToString("yyyy-MM-dd");
                                        }
                                        else
                                        {
                                            //value = DateTime.Parse(value.ToString());
                                            if (DateTime.TryParse(valueToImport.ToString(), out date))
                                            {
                                                valueToImport = date.ToString("yyyy-MM-dd");
                                            }
                                            else
                                            {
                                                valueToImport = DateTime.Now.ToString("yyyy-MM-dd");
                                            }
                                        }
                                        break;
                                    case "datetime":
                                        DateTime datetime;
                                        if (valueToImport == null)
                                        {
                                            valueToImport = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        }
                                        else
                                        {
                                            if (DateTime.TryParse(valueToImport.ToString(), out datetime))
                                            {
                                                valueToImport = datetime.ToString("yyyy-MM-dd HH:mm:ss");
                                            }
                                            else
                                            {
                                                valueToImport = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                            }
                                        }
                                        break;
                                    default:
                                        valueToImport = String.Format("{0}", valueToImport);
                                        break;
                                }
                                colInsertNames += String.Format("{0},", fieldName);
                                colInsertValues += String.Format("@{0},", fieldName);
                                fields.Add(new Field(fieldName, valueToImport));
                            }
                            else
                            {
                                if (string.IsNullOrWhiteSpace(col.ColumnDefault) && !col.IsNullable)
                                {
                                    switch (dataType)
                                    {
                                        case "int":
                                            valueToImport = 0;
                                            break;
                                        case "bit":
                                            valueToImport = 0;
                                            break;
                                        case "date":
                                            valueToImport = DateTime.Now.ToString("yyyy-MM-dd");
                                            break;
                                        case "datetime":
                                            valueToImport = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                            break;
                                        default:
                                            valueToImport = String.Format("{0}", valueToImport);
                                            break;
                                    }
                                    colInsertNames += String.Format("{0},", fieldName);
                                    colInsertValues += String.Format("@{0},", fieldName);
                                    fields.Add(new Field(fieldName, valueToImport));
                                }
                            }

                        }
                        var sqlCommand = String.Format("INSERT INTO {0} ({1}) VALUES ({2})", tableName, colInsertNames.Remove(colInsertNames.Length - 1, 1), colInsertValues.Remove(colInsertValues.Length - 1, 1));
                        //sqlCommand = sqlCommand.Replace(@"\", @"\\");
                        //mariaDbContext.CloseConnection();

                        if (deleteOldData)
                        {
                            mariaDbContext.ExecuteNonQuery("DELETE FROM " + tableName);
                            deleteOldData = false;// Gán lại để lần sau không xóa nữa.
                        }
                        try
                        {
                            foreach (var field in fields)
                            {
                                var fieldName = String.Format("@{0}", field.FieldName);
                                mariaDbContext.SetParamValue(fieldName, field.FieldValue);
                            }
                            var result = mariaDbContext.ExecuteNonQuery(sqlCommand);
                            _numberImportSuccess += result;
                            txtQuery.Invoke(new System.Action(() => txtQuery.AppendText(String.Format("{0}. [SUCCESS] INSERT {1} new record success. \r\n", _progessCount + 1, result))));
                        }
                        catch (Exception ex)
                        {
                            var countErro = rowsErro.Count;
                            var row = new Row(countErro);
                            for (int k = 1; k <= _excelColCount; k++)
                            {
                                var cell = new Cell();
                                cell.RowIndex = rowsErro.Count + 1;
                                cell.ColumnIndex = k;
                                var value = _xlRange.Cells[i, k].Value;
                                if (value == null)
                                    value = string.Empty;
                                cell.Value = value;
                                row.Cells.Add(cell);
                            }
                            row.Cells.Add(new Cell(rowsErro.Count + 1, (int)_excelColCount + 1, ex.Message));
                            rowsErro.Add(row);
                            _numberImportErro++;
                            txtQuery.Invoke(new System.Action(() => txtQuery.AppendText(String.Format("{0}.[FAIL] INSERT {1} new record.\r\n", _progessCount + 1, ex.Message))));
                        }

                        // Cập nhật tiến trình hoàn thành:
                        _progessCount++;
                        var pecentSuccess = (_numberImportSuccess * 100) / totalRecord;
                        var pecentErro = (_numberImportErro * 100) / totalRecord;
                        var totalTimes = watch.Elapsed.TotalSeconds;
                        var totalMinutes = (int)(totalTimes / 60);
                        var second = (int)(totalTimes % 60);
                        lblImportSuccessInfo.Invoke(new System.Action(() => lblImportSuccessInfo.Text = String.Format("{0}/{1} ({2}%)", _numberImportSuccess, totalRecord, pecentSuccess)));
                        lblTimeProgessInfo.Invoke(new System.Action(() => lblTimeProgessInfo.Text = String.Format("Time progess: {0}(m) {1}(s)", totalMinutes, second)));
                        if (_numberImportErro > 0)
                        {
                            lblImportErrorInfo.Invoke(new System.Action(() => lblImportErrorInfo.Text = String.Format("{0} ({1}%)", _numberImportErro, pecentErro)));
                        }
                        if (_worker.CancellationPending == true)
                        {
                            e.Cancel = true;
                            return;
                        }
                        var pecentProgess = (_progessCount * 100) / totalProgessCount;
                        var pecentProgessDecimal = Math.Round(((decimal)(_progessCount * 100) / (decimal)totalProgessCount), 2);
                        lblProgessInfo.Invoke(new System.Action(() => lblProgessInfo.Text = String.Format("Đang {0} ... ({1}%) - đã xử lý {2}/{3} bản ghi", _processInfo, pecentProgessDecimal, _progessCount, _excelRowCount - 1)));
                        _worker.ReportProgress(pecentProgess);
                    }
                }

                //--------------------------------------------------------------------------------------------------------------
                //cleanup  

                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                // Cập nhật tiến trình hoàn thành:
                _progessCount++;
                _worker.ReportProgress((_progessCount * 100) / totalProgessCount);

                //rule of thumb for releasing com objects:  
                //  never use two dots, all COM objects must be referenced and released individually  
                //  ex: [somthing].[something].[something] is bad  

                //release com objects to fully kill excel process from running in the background  
                Marshal.ReleaseComObject(_xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                // Cập nhật tiến trình hoàn thành:
                _progessCount++;
                _worker.ReportProgress((_progessCount * 100) / totalProgessCount);

                //close and release  
                _xlWorkbook.Close();
                Marshal.ReleaseComObject(_xlWorkbook);

                // Cập nhật tiến trình hoàn thành:
                _progessCount++;
                _worker.ReportProgress((_progessCount * 100) / totalProgessCount);

                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                watch.Stop();

                // Cập nhật tiến trình hoàn thành:
                _progessCount++;
                _worker.ReportProgress((_progessCount * 100) / totalProgessCount);
            }
            catch (COMException)
            {
                MessageBox.Show("Vui lòng kiểm tra lại File nhập khẩu. \n File có thể đã bị hỏng / xóa / đổi tên hoặc đang được sử dụng bởi một chương trình khác.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void backgroundWorker_RunWorkerImportCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                progressBar.Value = 0;
                lblProgessInfo.Text = string.Empty;
            }
            else
            {
                if (_numberImportErro > 0)
                {
                    var confirmResult = MessageBox.Show(String.Format("Có {0} bản không thể nhập khẩu.Nhấn [OK để xem chi tiết].", _numberImportErro),
                                     "Xem tệp nhập khẩu lỗi!",
                                     MessageBoxButtons.YesNo);
                    if (confirmResult == DialogResult.Yes)
                    {
                        workerExportErroFile.RunWorkerAsync();
                    }
                }
                lblProgessInfo.Text = "Hoàn thành";
            }
            EnabledAllControlInForm(true);
        }

        #endregion

        #region ************************* EXPORT Erro File ****************************************
        private void workerExportErroFile_DoWork(object sender, DoWorkEventArgs e)
        {
            ResetInfoProgess();
            // Bắt đầu tính thời gian:
            var watch = new System.Diagnostics.Stopwatch();
            watch.Start();

            _processInfo = "xuất khẩu";
            _progessCount = 0;
            _worker = sender as BackgroundWorker;
            // creating Excel Application  
            _Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            _Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            _Worksheet worksheet = null;

            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = "Exported Erro Data";
            foreach (var item in rowsErro)
            {
                foreach (var cell in item.Cells)
                {
                    worksheet.Cells[cell.RowIndex, cell.ColumnIndex] = cell.Value;
                }
                _progessCount++;
                _worker.ReportProgress(_progessCount * 100 / rowsErro.Count);

                var totalTimes = watch.Elapsed.TotalSeconds;
                var totalMinutes = (int)(totalTimes / 60);
                var second = (int)(totalTimes % 60);
                lblTimeProgessInfo.Invoke(new System.Action(() => lblTimeProgessInfo.Text = String.Format("Time progess: {0}(m) {1}(s)", totalMinutes, second)));
                if (_worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    return;
                }
            }

            // see the excel sheet behind the program  
            app.Visible = true;
        }

        private void workerExportErroFile_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            if (_progessCount <= _numberImportErro)
            {
                lblProgessInfo.Text = String.Format("Đang {0} ... ({1}%) - đã ghi {2}/{3} bản ghi lỗi", _processInfo, e.ProgressPercentage, _progessCount - 1, _numberImportErro - 1);
            }
        }

        private void workerExportErroFile_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                progressBar.Value = 0;
                lblProgessInfo.Text = string.Empty;
            }
            else
            {
                lblProgessInfo.Text = "Hoàn thành";
            }
            EnabledAllControlInForm(true);
        }

        #endregion

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;


        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                progressBar.Value = 0;
                lblProgessInfo.Text = string.Empty;
                //lblImportInfo.Text = "Nhập khẩu xong";
            }
            else
            {
                lblProgessInfo.Text = "Hoàn thành";
                //lblImportInfo.Text = "Hoàn thành";
            }
            EnabledAllControlInForm(true);
        }

        private void btnResetForm_Click(object sender, EventArgs e)
        {
            ResetInfoProgess();
            //this.Close();
            //frmConnectInfo frmConnectInfo = new frmConnectInfo();
            //frmConnectInfo.Show();
            dgvColumnMapping.Rows.Clear();
            dgvExcelData.Rows.Clear();
            dgvExcelData.Columns.Clear();
            progressBar.Value = 0;
            txtFilePath.Text = string.Empty;
            var controls = panel1.Controls;
            foreach (var control in controls)
            {
                if (control is ComboBox)
                {
                    ((ComboBox)control).Items.Clear();
                    ((ComboBox)control).Enabled = false;
                }

            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Hủy nhập khẩu dữ liệu?", "Xác nhận", MessageBoxButtons.OKCancel);
            if (result == DialogResult.OK)
            {
                _worker.CancelAsync();
            }
        }

        /// <summary>
        /// Enabled toàn bộ các control nhập liệu trong Form
        /// </summary>
        /// <param name="enabled"></param>
        private void EnabledAllControlInForm(bool enabled)
        {
            foreach (Control control in groupBoxServerInfo.Controls)
            {
                if (control is ComboBox)
                    ((ComboBox)control).Enabled = enabled;
                if (control is System.Windows.Forms.TextBox)
                    ((System.Windows.Forms.TextBox)control).Enabled = enabled;
                if (control is System.Windows.Forms.CheckBox)
                    ((System.Windows.Forms.CheckBox)control).Enabled = enabled;
                if (control is System.Windows.Forms.Button)
                    ((System.Windows.Forms.Button)control).Enabled = enabled;
            }
            foreach (Control control in groupBoxImportInfo.Controls)
            {
                if (control is ComboBox)
                    ((ComboBox)control).Enabled = false;
                if (control is System.Windows.Forms.TextBox)
                    ((System.Windows.Forms.TextBox)control).Enabled = enabled;
                if (control is System.Windows.Forms.CheckBox)
                    ((System.Windows.Forms.CheckBox)control).Enabled = enabled;
                if (control is System.Windows.Forms.Button)
                    ((System.Windows.Forms.Button)control).Enabled = enabled;
            }
            foreach (Control control in panel1.Controls)
            {
                if (control is ComboBox)
                    ((ComboBox)control).Enabled = false;
                if (control is System.Windows.Forms.TextBox)
                    ((System.Windows.Forms.TextBox)control).Enabled = enabled;
                if (control is System.Windows.Forms.CheckBox)
                    ((System.Windows.Forms.CheckBox)control).Enabled = enabled;
                if (control is System.Windows.Forms.Button)
                    ((System.Windows.Forms.Button)control).Enabled = enabled;
            }
        }

        private void ResetInfoProgess()
        {
            progressBar.Maximum = 100;
            progressBar.Step = 1;
            progressBar.Value = 0;
            lblImportCount.Text = string.Empty;
            lblImportSuccessInfo.Text = "0";
            lblImportErrorInfo.Text = "0";
            lblTimeProgessInfo.Text = string.Empty;
            //lblImportTime.Text = string.Empty;
        }
    }
}
