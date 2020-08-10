namespace MNV.ExcelDataToDB
{
    partial class FrmConnectInfo
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmConnectInfo));
            this.label1 = new System.Windows.Forms.Label();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnChooseFile = new System.Windows.Forms.Button();
            this.groupBoxServerInfo = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblTotalColumn = new System.Windows.Forms.Label();
            this.cbxTableList = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.cbxDatabases = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.cbxDbType = new System.Windows.Forms.ComboBox();
            this.txtConnect = new System.Windows.Forms.Button();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtPort = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtServerInfo = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnImport = new System.Windows.Forms.Button();
            this.groupBoxImportInfo = new System.Windows.Forms.GroupBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.cboRemoveOldData = new System.Windows.Forms.CheckBox();
            this.numColumnHeader = new System.Windows.Forms.NumericUpDown();
            this.label12 = new System.Windows.Forms.Label();
            this.numRowEnd = new System.Windows.Forms.NumericUpDown();
            this.numRowBegin = new System.Windows.Forms.NumericUpDown();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.btnResetForm = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.dgvExcelData = new System.Windows.Forms.DataGridView();
            this.dgvColumnMapping = new System.Windows.Forms.DataGridView();
            this.colDbFieldName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDataType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colExcelFieldName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colExcelFieldPossition = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.backgroundWorkerImport = new System.ComponentModel.BackgroundWorker();
            this.txtQuery = new System.Windows.Forms.TextBox();
            this.lblProgessInfo = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblImportCount = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.lblImportErrorInfo = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.lblImportSuccessInfo = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.lblTimeProgessInfo = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.workerExportErroFile = new System.ComponentModel.BackgroundWorker();
            this.groupBoxServerInfo.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBoxImportInfo.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColumnHeader)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numRowEnd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numRowBegin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvExcelData)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvColumnMapping)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Candara", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(13, 9);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(572, 39);
            this.label1.TabIndex = 0;
            this.label1.Text = "Import Data From Excel  - Mạnh Nguyễn ";
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(9, 37);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(428, 29);
            this.txtFilePath.TabIndex = 1;
            // 
            // btnChooseFile
            // 
            this.btnChooseFile.Location = new System.Drawing.Point(435, 37);
            this.btnChooseFile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnChooseFile.Name = "btnChooseFile";
            this.btnChooseFile.Size = new System.Drawing.Size(112, 29);
            this.btnChooseFile.TabIndex = 2;
            this.btnChooseFile.Text = "Duyệt tệp";
            this.btnChooseFile.UseVisualStyleBackColor = true;
            this.btnChooseFile.Click += new System.EventHandler(this.btnChooseFile_Click);
            // 
            // groupBoxServerInfo
            // 
            this.groupBoxServerInfo.Controls.Add(this.panel1);
            this.groupBoxServerInfo.Controls.Add(this.label9);
            this.groupBoxServerInfo.Controls.Add(this.cbxDbType);
            this.groupBoxServerInfo.Controls.Add(this.txtConnect);
            this.groupBoxServerInfo.Controls.Add(this.txtPassword);
            this.groupBoxServerInfo.Controls.Add(this.label5);
            this.groupBoxServerInfo.Controls.Add(this.txtUserName);
            this.groupBoxServerInfo.Controls.Add(this.label4);
            this.groupBoxServerInfo.Controls.Add(this.txtPort);
            this.groupBoxServerInfo.Controls.Add(this.label3);
            this.groupBoxServerInfo.Controls.Add(this.txtServerInfo);
            this.groupBoxServerInfo.Controls.Add(this.label2);
            this.groupBoxServerInfo.Location = new System.Drawing.Point(12, 66);
            this.groupBoxServerInfo.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBoxServerInfo.Name = "groupBoxServerInfo";
            this.groupBoxServerInfo.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBoxServerInfo.Size = new System.Drawing.Size(714, 359);
            this.groupBoxServerInfo.TabIndex = 3;
            this.groupBoxServerInfo.TabStop = false;
            this.groupBoxServerInfo.Text = "Thông tin kết nối";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.lblTotalColumn);
            this.panel1.Controls.Add(this.cbxTableList);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.cbxDatabases);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Location = new System.Drawing.Point(366, 37);
            this.panel1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(338, 285);
            this.panel1.TabIndex = 15;
            // 
            // lblTotalColumn
            // 
            this.lblTotalColumn.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotalColumn.Location = new System.Drawing.Point(118, 207);
            this.lblTotalColumn.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblTotalColumn.Name = "lblTotalColumn";
            this.lblTotalColumn.Size = new System.Drawing.Size(201, 37);
            this.lblTotalColumn.TabIndex = 15;
            this.lblTotalColumn.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cbxTableList
            // 
            this.cbxTableList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxTableList.Enabled = false;
            this.cbxTableList.FormattingEnabled = true;
            this.cbxTableList.Location = new System.Drawing.Point(15, 132);
            this.cbxTableList.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cbxTableList.Name = "cbxTableList";
            this.cbxTableList.Size = new System.Drawing.Size(302, 29);
            this.cbxTableList.TabIndex = 11;
            this.cbxTableList.SelectedIndexChanged += new System.EventHandler(this.cbxTableList_SelectedIndexChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(10, 215);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(103, 22);
            this.label11.TabIndex = 14;
            this.label11.Text = "Tổng số cột:";
            // 
            // cbxDatabases
            // 
            this.cbxDatabases.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxDatabases.Enabled = false;
            this.cbxDatabases.Font = new System.Drawing.Font("Candara", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxDatabases.ForeColor = System.Drawing.Color.Red;
            this.cbxDatabases.FormattingEnabled = true;
            this.cbxDatabases.Location = new System.Drawing.Point(15, 61);
            this.cbxDatabases.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cbxDatabases.Name = "cbxDatabases";
            this.cbxDatabases.Size = new System.Drawing.Size(302, 29);
            this.cbxDatabases.TabIndex = 4;
            this.cbxDatabases.SelectedIndexChanged += new System.EventHandler(this.cbxDatabases_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 36);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(131, 22);
            this.label6.TabIndex = 6;
            this.label6.Text = "Chọn Database:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(10, 107);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(230, 22);
            this.label10.TabIndex = 12;
            this.label10.Text = "Chọn bảng sẽ import dữ liệu:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(22, 42);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(75, 22);
            this.label9.TabIndex = 10;
            this.label9.Text = "DB type:";
            // 
            // cbxDbType
            // 
            this.cbxDbType.Enabled = false;
            this.cbxDbType.FormattingEnabled = true;
            this.cbxDbType.Items.AddRange(new object[] {
            "MySQL",
            "MariaDB",
            "MS SQL",
            "PostgreSQL"});
            this.cbxDbType.Location = new System.Drawing.Point(118, 37);
            this.cbxDbType.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cbxDbType.Name = "cbxDbType";
            this.cbxDbType.Size = new System.Drawing.Size(212, 29);
            this.cbxDbType.TabIndex = 9;
            this.cbxDbType.Text = "MariaDB";
            // 
            // txtConnect
            // 
            this.txtConnect.Font = new System.Drawing.Font("Candara", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtConnect.ForeColor = System.Drawing.Color.ForestGreen;
            this.txtConnect.Location = new System.Drawing.Point(118, 278);
            this.txtConnect.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtConnect.Name = "txtConnect";
            this.txtConnect.Size = new System.Drawing.Size(214, 45);
            this.txtConnect.TabIndex = 4;
            this.txtConnect.Text = "KẾT NỐI";
            this.txtConnect.UseVisualStyleBackColor = true;
            this.txtConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(118, 236);
            this.txtPassword.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(212, 29);
            this.txtPassword.TabIndex = 8;
            this.txtPassword.Text = "misa@2020";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(24, 241);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(94, 22);
            this.label5.TabIndex = 7;
            this.label5.Text = "Password: ";
            // 
            // txtUserName
            // 
            this.txtUserName.Location = new System.Drawing.Point(118, 194);
            this.txtUserName.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.Size = new System.Drawing.Size(212, 29);
            this.txtUserName.TabIndex = 6;
            this.txtUserName.Text = "root";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(24, 199);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(95, 22);
            this.label4.TabIndex = 5;
            this.label4.Text = "UserName:";
            // 
            // txtPort
            // 
            this.txtPort.Location = new System.Drawing.Point(118, 132);
            this.txtPort.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtPort.Name = "txtPort";
            this.txtPort.Size = new System.Drawing.Size(212, 29);
            this.txtPort.TabIndex = 3;
            this.txtPort.Text = "3306";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 137);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 22);
            this.label3.TabIndex = 2;
            this.label3.Text = "Port:";
            // 
            // txtServerInfo
            // 
            this.txtServerInfo.Location = new System.Drawing.Point(118, 90);
            this.txtServerInfo.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtServerInfo.Name = "txtServerInfo";
            this.txtServerInfo.Size = new System.Drawing.Size(212, 29);
            this.txtServerInfo.TabIndex = 1;
            this.txtServerInfo.Text = "192.168.15.18";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 95);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(63, 22);
            this.label2.TabIndex = 0;
            this.label2.Text = "Server:";
            // 
            // btnImport
            // 
            this.btnImport.Font = new System.Drawing.Font("Candara", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImport.ForeColor = System.Drawing.Color.Red;
            this.btnImport.Location = new System.Drawing.Point(18, 294);
            this.btnImport.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(529, 53);
            this.btnImport.TabIndex = 7;
            this.btnImport.Text = "IMPORT";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // groupBoxImportInfo
            // 
            this.groupBoxImportInfo.Controls.Add(this.checkBox1);
            this.groupBoxImportInfo.Controls.Add(this.cboRemoveOldData);
            this.groupBoxImportInfo.Controls.Add(this.numColumnHeader);
            this.groupBoxImportInfo.Controls.Add(this.label12);
            this.groupBoxImportInfo.Controls.Add(this.numRowEnd);
            this.groupBoxImportInfo.Controls.Add(this.numRowBegin);
            this.groupBoxImportInfo.Controls.Add(this.btnImport);
            this.groupBoxImportInfo.Controls.Add(this.label7);
            this.groupBoxImportInfo.Controls.Add(this.btnChooseFile);
            this.groupBoxImportInfo.Controls.Add(this.label8);
            this.groupBoxImportInfo.Controls.Add(this.txtFilePath);
            this.groupBoxImportInfo.Location = new System.Drawing.Point(735, 66);
            this.groupBoxImportInfo.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBoxImportInfo.Name = "groupBoxImportInfo";
            this.groupBoxImportInfo.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBoxImportInfo.Size = new System.Drawing.Size(556, 359);
            this.groupBoxImportInfo.TabIndex = 8;
            this.groupBoxImportInfo.TabStop = false;
            this.groupBoxImportInfo.Text = "Thông tin nhập khẩu";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(18, 260);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(228, 26);
            this.checkBox1.TabIndex = 16;
            this.checkBox1.Text = "Tự động nhận diện dữ liệu";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // cboRemoveOldData
            // 
            this.cboRemoveOldData.AutoSize = true;
            this.cboRemoveOldData.Location = new System.Drawing.Point(252, 260);
            this.cboRemoveOldData.Name = "cboRemoveOldData";
            this.cboRemoveOldData.Size = new System.Drawing.Size(199, 26);
            this.cboRemoveOldData.TabIndex = 15;
            this.cboRemoveOldData.Text = "Xóa toàn bộ dữ liệu cũ";
            this.cboRemoveOldData.UseVisualStyleBackColor = true;
            // 
            // numColumnHeader
            // 
            this.numColumnHeader.Location = new System.Drawing.Point(235, 103);
            this.numColumnHeader.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.numColumnHeader.Name = "numColumnHeader";
            this.numColumnHeader.Size = new System.Drawing.Size(312, 29);
            this.numColumnHeader.TabIndex = 11;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(8, 106);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(113, 22);
            this.label12.TabIndex = 10;
            this.label12.Text = "Dòng tiêu đề:";
            // 
            // numRowEnd
            // 
            this.numRowEnd.Location = new System.Drawing.Point(235, 187);
            this.numRowEnd.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.numRowEnd.Name = "numRowEnd";
            this.numRowEnd.Size = new System.Drawing.Size(312, 29);
            this.numRowEnd.TabIndex = 9;
            // 
            // numRowBegin
            // 
            this.numRowBegin.Location = new System.Drawing.Point(235, 145);
            this.numRowBegin.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.numRowBegin.Name = "numRowBegin";
            this.numRowBegin.Size = new System.Drawing.Size(312, 29);
            this.numRowBegin.TabIndex = 8;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(8, 190);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(176, 22);
            this.label7.TabIndex = 7;
            this.label7.Text = "Dòng dừng nhập liệu:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(8, 148);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(193, 22);
            this.label8.TabIndex = 5;
            this.label8.Text = "Dòng bắt đầu nhập liệu:";
            // 
            // btnResetForm
            // 
            this.btnResetForm.Location = new System.Drawing.Point(1147, 14);
            this.btnResetForm.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnResetForm.Name = "btnResetForm";
            this.btnResetForm.Size = new System.Drawing.Size(135, 53);
            this.btnResetForm.TabIndex = 14;
            this.btnResetForm.Text = "Reset";
            this.btnResetForm.UseVisualStyleBackColor = true;
            this.btnResetForm.Click += new System.EventHandler(this.btnResetForm_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 435);
            this.progressBar.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(1160, 33);
            this.progressBar.TabIndex = 12;
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            // 
            // dgvExcelData
            // 
            this.dgvExcelData.AllowUserToDeleteRows = false;
            this.dgvExcelData.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Candara", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvExcelData.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvExcelData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvExcelData.Location = new System.Drawing.Point(12, 499);
            this.dgvExcelData.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dgvExcelData.Name = "dgvExcelData";
            this.dgvExcelData.RowHeadersVisible = false;
            this.dgvExcelData.Size = new System.Drawing.Size(1276, 90);
            this.dgvExcelData.TabIndex = 9;
            // 
            // dgvColumnMapping
            // 
            this.dgvColumnMapping.AllowUserToAddRows = false;
            this.dgvColumnMapping.AllowUserToResizeRows = false;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.ActiveCaption;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Candara", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.MidnightBlue;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvColumnMapping.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvColumnMapping.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvColumnMapping.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colDbFieldName,
            this.colDataType,
            this.colExcelFieldName,
            this.colExcelFieldPossition});
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Candara", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.Padding = new System.Windows.Forms.Padding(6, 2, 6, 2);
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvColumnMapping.DefaultCellStyle = dataGridViewCellStyle4;
            this.dgvColumnMapping.GridColor = System.Drawing.SystemColors.ButtonFace;
            this.dgvColumnMapping.Location = new System.Drawing.Point(12, 621);
            this.dgvColumnMapping.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dgvColumnMapping.Name = "dgvColumnMapping";
            this.dgvColumnMapping.RowHeadersVisible = false;
            this.dgvColumnMapping.RowTemplate.Height = 26;
            this.dgvColumnMapping.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvColumnMapping.Size = new System.Drawing.Size(1275, 240);
            this.dgvColumnMapping.TabIndex = 10;
            // 
            // colDbFieldName
            // 
            this.colDbFieldName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colDbFieldName.DataPropertyName = "ColumnName";
            this.colDbFieldName.FillWeight = 150F;
            this.colDbFieldName.HeaderText = "DB Field";
            this.colDbFieldName.Name = "colDbFieldName";
            this.colDbFieldName.ReadOnly = true;
            // 
            // colDataType
            // 
            this.colDataType.DataPropertyName = "DataType";
            this.colDataType.HeaderText = "Data Type";
            this.colDataType.Name = "colDataType";
            this.colDataType.ReadOnly = true;
            this.colDataType.ToolTipText = "Kiểu dữ liệu";
            this.colDataType.Width = 150;
            // 
            // colExcelFieldName
            // 
            this.colExcelFieldName.DataPropertyName = "ExcelColHeaderName";
            this.colExcelFieldName.HeaderText = "Column In Excel";
            this.colExcelFieldName.Name = "colExcelFieldName";
            this.colExcelFieldName.ReadOnly = true;
            this.colExcelFieldName.Width = 170;
            // 
            // colExcelFieldPossition
            // 
            this.colExcelFieldPossition.DataPropertyName = "ExcelColPossition";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle3.ForeColor = System.Drawing.Color.Red;
            this.colExcelFieldPossition.DefaultCellStyle = dataGridViewCellStyle3;
            this.colExcelFieldPossition.HeaderText = "Excel Field Possition";
            this.colExcelFieldPossition.Name = "colExcelFieldPossition";
            this.colExcelFieldPossition.Width = 200;
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.WorkerReportsProgress = true;
            this.backgroundWorker.WorkerSupportsCancellation = true;
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BackgroundWorker_DoWorkReadData);
            this.backgroundWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            this.backgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerCompleted);
            // 
            // backgroundWorkerImport
            // 
            this.backgroundWorkerImport.WorkerReportsProgress = true;
            this.backgroundWorkerImport.WorkerSupportsCancellation = true;
            this.backgroundWorkerImport.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoImport);
            this.backgroundWorkerImport.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            this.backgroundWorkerImport.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerImportCompleted);
            // 
            // txtQuery
            // 
            this.txtQuery.Location = new System.Drawing.Point(12, 871);
            this.txtQuery.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtQuery.Multiline = true;
            this.txtQuery.Name = "txtQuery";
            this.txtQuery.ReadOnly = true;
            this.txtQuery.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtQuery.Size = new System.Drawing.Size(1278, 152);
            this.txtQuery.TabIndex = 11;
            // 
            // lblProgessInfo
            // 
            this.lblProgessInfo.AutoSize = true;
            this.lblProgessInfo.Font = new System.Drawing.Font("Candara", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblProgessInfo.ForeColor = System.Drawing.Color.ForestGreen;
            this.lblProgessInfo.Location = new System.Drawing.Point(522, 440);
            this.lblProgessInfo.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblProgessInfo.Name = "lblProgessInfo";
            this.lblProgessInfo.Size = new System.Drawing.Size(199, 22);
            this.lblProgessInfo.TabIndex = 13;
            this.lblProgessInfo.Text = "Đang nhập khẩu ... (20%)";
            this.lblProgessInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(1180, 434);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(112, 35);
            this.btnCancel.TabIndex = 14;
            this.btnCancel.Text = "Hủy";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblImportCount
            // 
            this.lblImportCount.AutoSize = true;
            this.lblImportCount.Font = new System.Drawing.Font("Candara", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblImportCount.ForeColor = System.Drawing.Color.ForestGreen;
            this.lblImportCount.Location = new System.Drawing.Point(718, 440);
            this.lblImportCount.Name = "lblImportCount";
            this.lblImportCount.Size = new System.Drawing.Size(139, 22);
            this.lblImportCount.TabIndex = 15;
            this.lblImportCount.Text = "thông tin import";
            this.lblImportCount.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Candara", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(1048, 1032);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(243, 17);
            this.label13.TabIndex = 16;
            this.label13.Text = "ExcelToDb - Beta Version ® manhnv.net";
            // 
            // lblImportErrorInfo
            // 
            this.lblImportErrorInfo.AutoSize = true;
            this.lblImportErrorInfo.Font = new System.Drawing.Font("Candara", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblImportErrorInfo.ForeColor = System.Drawing.Color.Red;
            this.lblImportErrorInfo.Location = new System.Drawing.Point(580, 1028);
            this.lblImportErrorInfo.Name = "lblImportErrorInfo";
            this.lblImportErrorInfo.Size = new System.Drawing.Size(20, 22);
            this.lblImportErrorInfo.TabIndex = 4;
            this.lblImportErrorInfo.Text = "0";
            this.lblImportErrorInfo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.ForeColor = System.Drawing.Color.Red;
            this.label17.Location = new System.Drawing.Point(532, 1029);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(52, 22);
            this.label17.TabIndex = 3;
            this.label17.Text = "Error:";
            // 
            // lblImportSuccessInfo
            // 
            this.lblImportSuccessInfo.AutoSize = true;
            this.lblImportSuccessInfo.Font = new System.Drawing.Font("Candara", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblImportSuccessInfo.ForeColor = System.Drawing.Color.ForestGreen;
            this.lblImportSuccessInfo.Location = new System.Drawing.Point(326, 1027);
            this.lblImportSuccessInfo.Name = "lblImportSuccessInfo";
            this.lblImportSuccessInfo.Size = new System.Drawing.Size(114, 22);
            this.lblImportSuccessInfo.TabIndex = 2;
            this.lblImportSuccessInfo.Text = "120/200 (60%)";
            this.lblImportSuccessInfo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.ForeColor = System.Drawing.Color.ForestGreen;
            this.label15.Location = new System.Drawing.Point(241, 1029);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(79, 22);
            this.label15.TabIndex = 1;
            this.label15.Text = "Success: ";
            // 
            // lblTimeProgessInfo
            // 
            this.lblTimeProgessInfo.AutoSize = true;
            this.lblTimeProgessInfo.ForeColor = System.Drawing.Color.Red;
            this.lblTimeProgessInfo.Location = new System.Drawing.Point(16, 1028);
            this.lblTimeProgessInfo.Name = "lblTimeProgessInfo";
            this.lblTimeProgessInfo.Size = new System.Drawing.Size(206, 22);
            this.lblTimeProgessInfo.TabIndex = 19;
            this.lblTimeProgessInfo.Text = "Time progess: 10(m) 20(s)";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(13, 472);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(226, 22);
            this.label14.TabIndex = 20;
            this.label14.Text = "Các cột dữ liệu trên file Excel";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(13, 594);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(293, 22);
            this.label16.TabIndex = 21;
            this.label16.Text = "Thổng tin bảng dữ liệu sẽ nhập khẩu:";
            // 
            // workerExportErroFile
            // 
            this.workerExportErroFile.WorkerReportsProgress = true;
            this.workerExportErroFile.WorkerSupportsCancellation = true;
            this.workerExportErroFile.DoWork += new System.ComponentModel.DoWorkEventHandler(this.workerExportErroFile_DoWork);
            this.workerExportErroFile.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.workerExportErroFile_ProgressChanged);
            this.workerExportErroFile.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.workerExportErroFile_RunWorkerCompleted);
            // 
            // FrmConnectInfo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1301, 1053);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.btnResetForm);
            this.Controls.Add(this.lblTimeProgessInfo);
            this.Controls.Add(this.lblImportErrorInfo);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.lblImportSuccessInfo);
            this.Controls.Add(this.lblImportCount);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblProgessInfo);
            this.Controls.Add(this.txtQuery);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.dgvColumnMapping);
            this.Controls.Add(this.dgvExcelData);
            this.Controls.Add(this.groupBoxImportInfo);
            this.Controls.Add(this.groupBoxServerInfo);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Candara", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FrmConnectInfo";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import Data From Excel";
            this.Load += new System.EventHandler(this.frmConnectInfo_Load);
            this.groupBoxServerInfo.ResumeLayout(false);
            this.groupBoxServerInfo.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBoxImportInfo.ResumeLayout(false);
            this.groupBoxImportInfo.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColumnHeader)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numRowEnd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numRowBegin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvExcelData)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvColumnMapping)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnChooseFile;
        private System.Windows.Forms.GroupBox groupBoxServerInfo;
        private System.Windows.Forms.TextBox txtPort;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtServerInfo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button txtConnect;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtUserName;
        private System.Windows.Forms.ComboBox cbxDatabases;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.GroupBox groupBoxImportInfo;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.NumericUpDown numRowEnd;
        private System.Windows.Forms.NumericUpDown numRowBegin;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cbxDbType;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.DataGridView dgvExcelData;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox cbxTableList;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.DataGridView dgvColumnMapping;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.NumericUpDown numColumnHeader;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button btnResetForm;
        private System.Windows.Forms.TextBox txtQuery;
        public System.ComponentModel.BackgroundWorker backgroundWorkerImport;
        public System.ComponentModel.BackgroundWorker backgroundWorker;
        public System.Windows.Forms.Label lblProgessInfo;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblTotalColumn;
        private System.Windows.Forms.CheckBox cboRemoveOldData;
        private System.Windows.Forms.Label lblImportCount;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label lblImportErrorInfo;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label lblImportSuccessInfo;
        private System.Windows.Forms.Label lblTimeProgessInfo;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDbFieldName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDataType;
        private System.Windows.Forms.DataGridViewTextBoxColumn colExcelFieldName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colExcelFieldPossition;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label16;
        public System.ComponentModel.BackgroundWorker workerExportErroFile;
    }
}

