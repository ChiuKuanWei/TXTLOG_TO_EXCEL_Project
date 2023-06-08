namespace TXTLOG_TO_EXCEL_Project
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.lbl_MDB = new System.Windows.Forms.Label();
            this.txt_MDBPATH = new System.Windows.Forms.TextBox();
            this.dlg_MDB = new System.Windows.Forms.OpenFileDialog();
            this.txt_TXTPATH = new System.Windows.Forms.TextBox();
            this.lbl_TXT = new System.Windows.Forms.Label();
            this.dlg_TXT = new System.Windows.Forms.OpenFileDialog();
            this.btn_Report = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lbl_MDB
            // 
            this.lbl_MDB.AutoSize = true;
            this.lbl_MDB.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_MDB.Location = new System.Drawing.Point(21, 26);
            this.lbl_MDB.Name = "lbl_MDB";
            this.lbl_MDB.Size = new System.Drawing.Size(89, 16);
            this.lbl_MDB.TabIndex = 0;
            this.lbl_MDB.Text = "MDB檔選擇";
            this.lbl_MDB.Click += new System.EventHandler(this.lbl_MDB_Click);
            // 
            // txt_MDBPATH
            // 
            this.txt_MDBPATH.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_MDBPATH.Location = new System.Drawing.Point(119, 23);
            this.txt_MDBPATH.Name = "txt_MDBPATH";
            this.txt_MDBPATH.ReadOnly = true;
            this.txt_MDBPATH.Size = new System.Drawing.Size(586, 22);
            this.txt_MDBPATH.TabIndex = 1;
            this.txt_MDBPATH.TabStop = false;
            // 
            // dlg_MDB
            // 
            this.dlg_MDB.FileName = "openFileDialog1";
            // 
            // txt_TXTPATH
            // 
            this.txt_TXTPATH.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_TXTPATH.Location = new System.Drawing.Point(119, 54);
            this.txt_TXTPATH.Name = "txt_TXTPATH";
            this.txt_TXTPATH.ReadOnly = true;
            this.txt_TXTPATH.Size = new System.Drawing.Size(586, 27);
            this.txt_TXTPATH.TabIndex = 3;
            this.txt_TXTPATH.TabStop = false;
            // 
            // lbl_TXT
            // 
            this.lbl_TXT.AutoSize = true;
            this.lbl_TXT.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_TXT.Location = new System.Drawing.Point(21, 57);
            this.lbl_TXT.Name = "lbl_TXT";
            this.lbl_TXT.Size = new System.Drawing.Size(86, 16);
            this.lbl_TXT.TabIndex = 2;
            this.lbl_TXT.Text = "LOG檔選擇";
            this.lbl_TXT.Click += new System.EventHandler(this.lbl_TXT_Click);
            // 
            // dlg_TXT
            // 
            this.dlg_TXT.FileName = "openFileDialog1";
            // 
            // btn_Report
            // 
            this.btn_Report.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Report.Location = new System.Drawing.Point(711, 55);
            this.btn_Report.Name = "btn_Report";
            this.btn_Report.Size = new System.Drawing.Size(72, 26);
            this.btn_Report.TabIndex = 4;
            this.btn_Report.Text = "轉檔";
            this.btn_Report.UseVisualStyleBackColor = true;
            this.btn_Report.Click += new System.EventHandler(this.btn_Report_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 97);
            this.Controls.Add(this.btn_Report);
            this.Controls.Add(this.txt_TXTPATH);
            this.Controls.Add(this.lbl_TXT);
            this.Controls.Add(this.txt_MDBPATH);
            this.Controls.Add(this.lbl_MDB);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_MDB;
        private System.Windows.Forms.TextBox txt_MDBPATH;
        private System.Windows.Forms.OpenFileDialog dlg_MDB;
        private System.Windows.Forms.TextBox txt_TXTPATH;
        private System.Windows.Forms.Label lbl_TXT;
        private System.Windows.Forms.OpenFileDialog dlg_TXT;
        private System.Windows.Forms.Button btn_Report;
    }
}

