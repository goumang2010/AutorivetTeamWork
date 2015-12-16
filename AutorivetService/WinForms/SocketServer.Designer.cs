namespace AutorivetService
{
    partial class frm_server
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
            this.txtIp = new System.Windows.Forms.ComboBox();
            this.txtPort = new System.Windows.Forms.ComboBox();
            this.lbOnline = new System.Windows.Forms.ListBox();
            this.txtMsg = new System.Windows.Forms.TextBox();
            this.txtMsgSend = new System.Windows.Forms.TextBox();
            this.txtSelectFile = new System.Windows.Forms.TextBox();
            this.btnSend = new System.Windows.Forms.Button();
            this.btnSendToAll = new System.Windows.Forms.Button();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.btnSendFile = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtIp
            // 
            this.txtIp.FormattingEnabled = true;
            this.txtIp.Location = new System.Drawing.Point(60, 71);
            this.txtIp.Name = "txtIp";
            this.txtIp.Size = new System.Drawing.Size(121, 20);
            this.txtIp.TabIndex = 0;
            this.txtIp.Text = "192.168.3.32";
            // 
            // txtPort
            // 
            this.txtPort.FormattingEnabled = true;
            this.txtPort.Location = new System.Drawing.Point(228, 71);
            this.txtPort.Name = "txtPort";
            this.txtPort.Size = new System.Drawing.Size(121, 20);
            this.txtPort.TabIndex = 1;
            this.txtPort.Text = "88";
            // 
            // lbOnline
            // 
            this.lbOnline.FormattingEnabled = true;
            this.lbOnline.ItemHeight = 12;
            this.lbOnline.Location = new System.Drawing.Point(376, 139);
            this.lbOnline.Name = "lbOnline";
            this.lbOnline.Size = new System.Drawing.Size(120, 88);
            this.lbOnline.TabIndex = 2;
            // 
            // txtMsg
            // 
            this.txtMsg.AcceptsReturn = true;
            this.txtMsg.Location = new System.Drawing.Point(96, 139);
            this.txtMsg.Multiline = true;
            this.txtMsg.Name = "txtMsg";
            this.txtMsg.Size = new System.Drawing.Size(100, 160);
            this.txtMsg.TabIndex = 3;
            // 
            // txtMsgSend
            // 
            this.txtMsgSend.Location = new System.Drawing.Point(96, 323);
            this.txtMsgSend.Name = "txtMsgSend";
            this.txtMsgSend.Size = new System.Drawing.Size(100, 21);
            this.txtMsgSend.TabIndex = 4;
            // 
            // txtSelectFile
            // 
            this.txtSelectFile.Location = new System.Drawing.Point(376, 293);
            this.txtSelectFile.Name = "txtSelectFile";
            this.txtSelectFile.Size = new System.Drawing.Size(100, 21);
            this.txtSelectFile.TabIndex = 5;
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(96, 350);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(75, 23);
            this.btnSend.TabIndex = 7;
            this.btnSend.Text = "发送";
            this.btnSend.UseVisualStyleBackColor = true;
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // btnSendToAll
            // 
            this.btnSendToAll.Location = new System.Drawing.Point(237, 350);
            this.btnSendToAll.Name = "btnSendToAll";
            this.btnSendToAll.Size = new System.Drawing.Size(75, 23);
            this.btnSendToAll.TabIndex = 8;
            this.btnSendToAll.Text = "群发";
            this.btnSendToAll.UseVisualStyleBackColor = true;
            this.btnSendToAll.Click += new System.EventHandler(this.btnSendToAll_Click);
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(96, 398);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(75, 23);
            this.btnSelectFile.TabIndex = 9;
            this.btnSelectFile.Text = "选择文件";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click_1);
            // 
            // btnSendFile
            // 
            this.btnSendFile.Location = new System.Drawing.Point(237, 398);
            this.btnSendFile.Name = "btnSendFile";
            this.btnSendFile.Size = new System.Drawing.Size(75, 23);
            this.btnSendFile.TabIndex = 10;
            this.btnSendFile.Text = "发送文件";
            this.btnSendFile.UseVisualStyleBackColor = true;
            this.btnSendFile.Click += new System.EventHandler(this.btnSendFile_Click_1);
            // 
            // frm_server
            // 
            this.ClientSize = new System.Drawing.Size(552, 480);
            this.Controls.Add(this.btnSendFile);
            this.Controls.Add(this.btnSelectFile);
            this.Controls.Add(this.btnSendToAll);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.txtSelectFile);
            this.Controls.Add(this.txtMsgSend);
            this.Controls.Add(this.txtMsg);
            this.Controls.Add(this.lbOnline);
            this.Controls.Add(this.txtPort);
            this.Controls.Add(this.txtIp);
            this.Name = "frm_server";
            this.Load += new System.EventHandler(this.frm_server_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox txtIp;
        private System.Windows.Forms.ComboBox txtPort;
        private System.Windows.Forms.ListBox lbOnline;
        private System.Windows.Forms.TextBox txtMsg;
        private System.Windows.Forms.TextBox txtMsgSend;
        private System.Windows.Forms.TextBox txtSelectFile;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.Button btnSendToAll;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.Button btnSendFile;
    }
}