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
using OFFICE_Method;

namespace AUTORIVET_KAOHE
{
    public partial class frm_server : Form
    {
        public frm_server()
        {
            InitializeComponent();
            TextBox.CheckForIllegalCrossThreadCalls = false;
        }

        Thread threadWatch = null; // 负责监听客户端连接请求的 线程；
        Socket socketWatch = null;

        Dictionary<string, Socket> dict = new Dictionary<string, Socket>();
        Dictionary<string, Thread> dictThread = new Dictionary<string, Thread>();

        private void btnBeginListen_Click(object sender, EventArgs e)
        {
            // 创建负责监听的套接字，注意其中的参数；
            socketWatch = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            // 获得文本框中的IP对象；
            IPAddress address = IPAddress.Parse(txtIp.Text.Trim());
                // 创建包含ip和端口号的网络节点对象；
                IPEndPoint endPoint = new IPEndPoint(address, int.Parse(txtPort.Text.Trim()));
                try
                {
                    // 将负责监听的套接字绑定到唯一的ip和端口上；
                    socketWatch.Bind(endPoint);
                }
                catch (SocketException se)
                {
                    MessageBox.Show("异常："+se.Message);
                    return;
                }
                // 设置监听队列的长度；
                socketWatch.Listen(10);
                // 创建负责监听的线程；
                threadWatch = new Thread(WatchConnecting);
                threadWatch.IsBackground = true;
                threadWatch.Start();
                ShowMsg("服务器启动监听成功！");
            //}
        }

        /// <summary>
        /// 监听客户端请求的方法；
        /// </summary>
        void WatchConnecting()
        {
            while (true)  // 持续不断的监听客户端的连接请求；
            {
                // 开始监听客户端连接请求，Accept方法会阻断当前的线程；
                Socket sokConnection = socketWatch.Accept(); // 一旦监听到一个客户端的请求，就返回一个与该客户端通信的 套接字；
                // 想列表控件中添加客户端的IP信息；
                lbOnline.Items.Add(sokConnection.RemoteEndPoint.ToString());
                // 将与客户端连接的 套接字 对象添加到集合中；
                dict.Add(sokConnection.RemoteEndPoint.ToString(), sokConnection);
                ShowMsg("客户端连接成功！");
                Thread thr = new Thread(RecMsg);
                thr.IsBackground = true;
                thr.Start(sokConnection);
                dictThread.Add(sokConnection.RemoteEndPoint.ToString(), thr);  //  将新建的线程 添加 到线程的集合中去。
            }
        }

        void RecMsg(object sokConnectionparn)
        {
                Socket sokClient = sokConnectionparn as Socket;
                while (true)
                {
                    // 定义一个2M的缓存区；
                    byte[] arrMsgRec = new byte[1024 * 1024 * 2];
                    // 将接受到的数据存入到输入  arrMsgRec中；
                    int length = -1;
                    try
                    {
                        length = sokClient.Receive(arrMsgRec); // 接收数据，并返回数据的长度；
                    }
                    catch (SocketException se)
                    {
                        ShowMsg("异常：" + se.Message);
                        // 从 通信套接字 集合中删除被中断连接的通信套接字；
                        dict.Remove(sokClient.RemoteEndPoint.ToString());
                        // 从通信线程集合中删除被中断连接的通信线程对象；
                        dictThread.Remove(sokClient.RemoteEndPoint.ToString());
                        // 从列表中移除被中断的连接IP
                        lbOnline.Items.Remove(sokClient.RemoteEndPoint.ToString());
                        break;
                    }
                    catch (Exception e)
                    {
                        ShowMsg("异常：" + e.Message);
                        // 从 通信套接字 集合中删除被中断连接的通信套接字；
                        dict.Remove(sokClient.RemoteEndPoint.ToString());
                        // 从通信线程集合中删除被中断连接的通信线程对象；
                        dictThread.Remove(sokClient.RemoteEndPoint.ToString());
                        // 从列表中移除被中断的连接IP
                        lbOnline.Items.Remove(sokClient.RemoteEndPoint.ToString());
                        break;
                    }
                    if (arrMsgRec[0] == 0)  // 表示接收到的是数据；
                    {
                        if(length==0)
                        {
                            // 从 通信套接字 集合中删除被中断连接的通信套接字；
                            dict.Remove(sokClient.RemoteEndPoint.ToString());
                            // 从通信线程集合中删除被中断连接的通信线程对象；
                            dictThread.Remove(sokClient.RemoteEndPoint.ToString());
                            // 从列表中移除被中断的连接IP
                            lbOnline.Items.Remove(sokClient.RemoteEndPoint.ToString());
                            break;
                        }
                        else
                        {
                            string strMsg = System.Text.Encoding.UTF8.GetString(arrMsgRec, 1, length - 1);// 将接受到的字节数据转化成字符串；
                            //if (strMsg.Contains("output_RNC"))
                            //{
                            //    string target=@"\\192.168.3.32\Autorivet\output\INFO\backup\RNC总表.xlsx";
                            //    excelMethod.SaveDataTableToExcel(DbHelperSQL.Query("select * from rnc总表").Tables[0], target);
                            //    System.Threading.Thread.Sleep(3000);
                            //    sendfile(target, sokClient.RemoteEndPoint.ToString());
                            //}
                            if (strMsg.Contains("notify:"))
                        {

                            notify(strMsg.Replace("notify:",""));
                        }
                            ShowMsg(strMsg);
                        }
                      
                    }
                    if (arrMsgRec[0] == 1) // 表示接收到的是文件；
                    {
                            SaveFileDialog sfd = new SaveFileDialog();
                           
                            if (sfd.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                            {// 在上边的 sfd.ShowDialog（） 的括号里边一定要加上 this 否则就不会弹出 另存为 的对话框，而弹出的是本类的其他窗口，，这个一定要注意！！！【解释：加了this的sfd.ShowDialog(this)，“另存为”窗口的指针才能被SaveFileDialog的对象调用，若不加thisSaveFileDialog 的对象调用的是本类的其他窗口了，当然不弹出“另存为”窗口。】
                               
                                string fileSavePath = sfd.FileName;// 获得文件保存的路径；
                                // 创建文件流，然后根据路径创建文件；
                                using (FileStream fs = new FileStream(fileSavePath, FileMode.Create))
                                {
                                    fs.Write(arrMsgRec, 1, length - 1);
                                    ShowMsg("文件保存成功：" + fileSavePath);
                                }
                            }
                        }
                }     
        }

        void ShowMsg(string str)
        {
           
            txtMsg.AppendText(str + "\r\n");
        }

        // 发送消息
        private void btnSend_Click(object sender, EventArgs e)
        {
            string strMsg = "服务器" + "\r\n" + "   -->" + txtMsgSend.Text.Trim() + "\r\n";
            byte[] arrMsg = System.Text.Encoding.UTF8.GetBytes(strMsg); // 将要发送的字符串转换成Utf-8字节数组；
            byte[] arrSendMsg=new byte[arrMsg.Length+1];
            arrSendMsg[0] = 0; // 表示发送的是消息数据
            Buffer.BlockCopy(arrMsg, 0, arrSendMsg, 1, arrMsg.Length);
            string strKey = "";
            strKey = lbOnline.Text.Trim();
            if (string.IsNullOrEmpty(strKey))   // 判断是不是选择了发送的对象；
            {
                MessageBox.Show("请选择你要发送的好友！！！");
            }
            else
            {
                dict[strKey].Send(arrSendMsg);// 解决了 sokConnection是局部变量，不能再本函数中引用的问题；
                ShowMsg(strMsg);
                txtMsgSend.Clear();
            }
        }

        /// <summary>
        /// 群发消息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">消息</param>
        private void btnSendToAll_Click(object sender, EventArgs e)
        {
            string strMsg = "服务器" + "\r\n" + "   -->" + txtMsgSend.Text.Trim() + "\r\n";
            byte[] arrMsg = System.Text.Encoding.UTF8.GetBytes(strMsg); // 将要发送的字符串转换成Utf-8字节数组；
       
           byte[] arrSendMsg = new byte[arrMsg.Length + 1]; // 上次写的时候把这一段给弄掉了，实在是抱歉哈~ 用来标识发送是数据而不是文件，如果没有这一段的客户端就接收不到消息了~~~
 arrSendMsg[0] = 0; // 表示发送的是消息数据
 Buffer.BlockCopy(arrMsg, 0, arrSendMsg, 1, arrMsg.Length);

                
            foreach (Socket s in dict.Values)
            {
              
                s.Send(arrMsg);
            }
            ShowMsg(strMsg);
            txtMsgSend.Clear();
            ShowMsg(" 群发完毕～～～");
        }

        //通知所有服务器IP的用户
        public void notify(string msg)
        {
            string strMsg ="show:"+ msg + "\r\n";
            byte[] arrMsg = System.Text.Encoding.UTF8.GetBytes(strMsg); // 将要发送的字符串转换成Utf-8字节数组；
            byte[] arrSendMsg = new byte[arrMsg.Length + 1];
            arrSendMsg[0] = 0; // 表示发送的是消息数据
            Buffer.BlockCopy(arrMsg, 0, arrSendMsg, 1, arrMsg.Length);

          foreach( var pp in lbOnline.Items)
            {
                string dicKey= pp.ToString();
                if(dicKey.Contains("192.168.3.32"))
                {
                    dict[dicKey].Send(arrSendMsg);
                    ShowMsg(strMsg);
                    txtMsgSend.Clear();
                }



            }




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
        private void sendfile(string filepath,string strKey)
        {
                            using (FileStream fs = new FileStream(filepath, FileMode.Open))
                {
                    string fileName=System.IO.Path.GetFileName(filepath);
                    string fileExtension=System.IO.Path.GetExtension(filepath);
                    string strMsg = "我给你发送的文件为:"+fileName+"\r\n";
                    byte[] arrMsg = System.Text.Encoding.UTF8.GetBytes(strMsg); // 将要发送的字符串转换成Utf-8字节数组；
                    byte[] arrSendMsg = new byte[arrMsg.Length + 1];
                    arrSendMsg[0] = 0; // 表示发送的是消息数据
                    Buffer.BlockCopy(arrMsg, 0, arrSendMsg, 1, arrMsg.Length);
                  //  bool fff = true;
                   
                    if (string.IsNullOrEmpty(strKey))   // 判断是不是选择了发送的对象；
                    {
                        MessageBox.Show("请选择你要发送的好友！！！");
                    }
                    else
                    {
                    dict[strKey].Send(arrSendMsg);// 解决了 sokConnection是局部变量，不能再本函数中引用的问题；
                    byte[] arrFile = new byte[1024 * 1024 * 2];
                    int length = fs.Read(arrFile, 0, arrFile.Length);  // 将文件中的数据读到arrFile数组中；
                    byte[] arrFileSend = new byte[length + 1];
                    arrFileSend[0] = 1; // 用来表示发送的是文件数据；
                    Buffer.BlockCopy(arrFile, 0, arrFileSend, 1, length);
                    // 还有一个 CopyTo的方法，但是在这里不适合； 当然还可以用for循环自己转化；
                    //  sockClient.Send(arrFileSend);// 发送数据到服务端；
                    dict[strKey].Send(arrFileSend);// 解决了 sokConnection是局部变量，不能再本函数中引用的问题；
                    
                    }
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
                sendfile(txtSelectFile.Text, strKey);
                txtSelectFile.Clear(); 
            }
            txtSelectFile.Clear();
        }
        private void frm_server_Load(object sender, EventArgs e)
        {

        }

        private void txtMsg_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

      
    