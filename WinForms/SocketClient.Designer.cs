﻿namespace AUTORIVET_KAOHE
{
    partial class frmClient
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
            this.txtPort = new System.Windows.Forms.ComboBox();
            this.txtIp = new System.Windows.Forms.ComboBox();
            this.txtMsg = new System.Windows.Forms.TextBox();
            this.btnSendFile = new System.Windows.Forms.Button();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.btnSend = new System.Windows.Forms.Button();
            this.btnBeginListen = new System.Windows.Forms.Button();
            this.txtSendMsg = new System.Windows.Forms.TextBox();
            this.txtSelectFile = new System.Windows.Forms.TextBox();
            this.txtName = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtPort
            // 
            this.txtPort.FormattingEnabled = true;
            this.txtPort.Location = new System.Drawing.Point(235, 69);
            this.txtPort.Name = "txtPort";
            this.txtPort.Size = new System.Drawing.Size(75, 20);
            this.txtPort.TabIndex = 3;
            this.txtPort.Text = "88";
            // 
            // txtIp
            // 
            this.txtIp.FormattingEnabled = true;
            this.txtIp.Location = new System.Drawing.Point(90, 69);
            this.txtIp.Name = "txtIp";
            this.txtIp.Size = new System.Drawing.Size(121, 20);
            this.txtIp.TabIndex = 2;
            this.txtIp.Text = "192.168.3.32";
            // 
            // txtMsg
            // 
            this.txtMsg.AcceptsReturn = true;
            this.txtMsg.Location = new System.Drawing.Point(88, 161);
            this.txtMsg.Multiline = true;
            this.txtMsg.Name = "txtMsg";
            this.txtMsg.Size = new System.Drawing.Size(100, 160);
            this.txtMsg.TabIndex = 4;
            // 
            // btnSendFile
            // 
            this.btnSendFile.Location = new System.Drawing.Point(235, 536);
            this.btnSendFile.Name = "btnSendFile";
            this.btnSendFile.Size = new System.Drawing.Size(75, 23);
            this.btnSendFile.TabIndex = 12;
            this.btnSendFile.Text = "发送文件";
            this.btnSendFile.UseVisualStyleBackColor = true;
            this.btnSendFile.Click += new System.EventHandler(this.btnSendFile_Click);
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(88, 460);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(75, 23);
            this.btnSelectFile.TabIndex = 11;
            this.btnSelectFile.Text = "选择文件";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(88, 413);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(75, 23);
            this.btnSend.TabIndex = 13;
            this.btnSend.Text = "发送";
            this.btnSend.UseVisualStyleBackColor = true;
            this.btnSend.Click += new System.EventHandler(this.btnSendMsg_Click);
            // 
            // btnBeginListen
            // 
            this.btnBeginListen.Location = new System.Drawing.Point(235, 161);
            this.btnBeginListen.Name = "btnBeginListen";
            this.btnBeginListen.Size = new System.Drawing.Size(75, 23);
            this.btnBeginListen.TabIndex = 14;
            this.btnBeginListen.Text = "连接";
            this.btnBeginListen.UseVisualStyleBackColor = true;
            this.btnBeginListen.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // txtSendMsg
            // 
            this.txtSendMsg.Location = new System.Drawing.Point(88, 363);
            this.txtSendMsg.Name = "txtSendMsg";
            this.txtSendMsg.Size = new System.Drawing.Size(100, 21);
            this.txtSendMsg.TabIndex = 15;
            // 
            // txtSelectFile
            // 
            this.txtSelectFile.Location = new System.Drawing.Point(88, 538);
            this.txtSelectFile.Name = "txtSelectFile";
            this.txtSelectFile.Size = new System.Drawing.Size(100, 21);
            this.txtSelectFile.TabIndex = 16;
            // 
            // txtName
            // 
            this.txtName.AutoSize = true;
            this.txtName.Location = new System.Drawing.Point(88, 328);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(41, 12);
            this.txtName.TabIndex = 18;
            this.txtName.Text = "label1";
            // 
            // frmClient
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(376, 616);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.txtSelectFile);
            this.Controls.Add(this.txtSendMsg);
            this.Controls.Add(this.btnBeginListen);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.btnSendFile);
            this.Controls.Add(this.btnSelectFile);
            this.Controls.Add(this.txtMsg);
            this.Controls.Add(this.txtPort);
            this.Controls.Add(this.txtIp);
            this.Name = "frmClient";
            this.Text = "frmClient";
            this.Load += new System.EventHandler(this.frmClient_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox txtPort;
        private System.Windows.Forms.ComboBox txtIp;
        private System.Windows.Forms.TextBox txtMsg;
        private System.Windows.Forms.Button btnSendFile;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.Button btnBeginListen;
        private System.Windows.Forms.TextBox txtSendMsg;
        private System.Windows.Forms.TextBox txtSelectFile;
        private System.Windows.Forms.Label txtName;
    }
}