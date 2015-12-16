using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Net;  // IP，IPAddress, IPEndPoint，端口等；
using System.Threading;
using System.IO;
using mysqlsolution;


namespace AutorivetService
{
    public partial class frm_server : Form
    {
        SocketServerService localServer;

        public frm_server()
        {

            InitializeComponent();
            localServer = new SocketServerService();
            lbOnline.DataBindings.Add("DataSource", localServer, "ClientList");

            //TextBox.CheckForIllegalCrossThreadCalls = false;
        }




        void ShowMsg(string str)
        {
           
            txtMsg.AppendText(str + "\r\n");
        }

        // 发送消息
        private void btnSend_Click(object sender, EventArgs e)
        {
            
            string strMsg = "服务器" + "\r\n" + "   -->" + txtMsgSend.Text.Trim() + "\r\n";
            localServer.SendMsg(lbOnline.SelectedItem.ToString(),strMsg);
            ShowMsg(strMsg);
            txtMsgSend.Clear();
            
        }

   
        // 群发消息
        
        private void btnSendToAll_Click(object sender, EventArgs e)
        {
            string strMsg = "服务器" + "\r\n" + "   -->" + txtMsgSend.Text.Trim() + "\r\n";
            localServer.SendMsgToAll(strMsg);
            ShowMsg(strMsg);
            txtMsgSend.Clear();
            ShowMsg("群发完毕!");
        }

      


        // 选择要发送的文件
        private void btnSelectFile_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = "D:\\";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtSelectFile.Text = ofd.FileName;
            }
        }
      
        // 文件的发送
        private void btnSendFile_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSelectFile.Text))
            {
                MessageBox.Show("请选择你要发送的文件！！！");
            }
            else
            {
                // 用文件流打开用户要发送的文件；
                string strKey = "";
                strKey = lbOnline.Text.Trim();
                localServer.SendFile(txtSelectFile.Text, strKey);
                txtSelectFile.Clear(); 
            }
            txtSelectFile.Clear();
        }

        private void frm_server_Load(object sender, EventArgs e)
        {
            try
            {
                localServer.btnBeginListen();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
                this.Close();
                return;
            }


            ShowMsg("服务器启动监听成功！");
        }
    }
}

      
    