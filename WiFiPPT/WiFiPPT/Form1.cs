using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Threading;
using System.Net;
using System.Net.Sockets;


namespace WiFiPPT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public bool servicesStatus = true;  //开始停止服务按钮状态
        public Thread myThread;       //声明一个线程实例
        public Socket newsock;        //声明一个Socket实例
        public Socket server;
        public Socket client;
        public EndPoint remote;
        public IPEndPoint localEP;
        public delegate void MyInvoke(string str);

        //监听函数
        public void Listen()
        {
            //初始化SOCKET实例
            newsock = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            //允许SOCKET被绑定在已使用的地址上。
            newsock.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            //初始化终结点实例
            localEP = new IPEndPoint(IPAddress.Any, int.Parse(portTextBox.Text.Trim()));
            try
            {
                //绑定
                newsock.Bind(localEP);
                //监听
                newsock.Listen(10);
                //开始接受连接，异步。
                newsock.BeginAccept(new AsyncCallback(OnConnectRequest), newsock);
            }
            catch (Exception ex)
            {
                //showClientMsg(ex.Message);
            }
        }

        //当有客户端连接时的处理
        public void OnConnectRequest(IAsyncResult ar)
        {
            //初始化一个SOCKET，用于其它客户端的连接
            server = (Socket)ar.AsyncState;
            client = server.EndAccept(ar);
            //将要发送给连接上来的客户端的提示字符串
            DateTimeOffset now = DateTimeOffset.Now;
            string strDateLine = "欢迎登录到服务器";
            Byte[] byteDateLine = System.Text.Encoding.UTF8.GetBytes(strDateLine);
            //将提示信息发送给客户端,并在服务端显示连接信息。
            remote = client.RemoteEndPoint;
            showClientMsg(client.RemoteEndPoint.ToString() + "连接成功。" + now.ToString("G") + "\r\n");
            client.Send(byteDateLine, byteDateLine.Length, 0);

            //等待新的客户端连接
            server.BeginAccept(new AsyncCallback(OnConnectRequest), server);

            while (true)
            {
                int recv = client.Receive(byteDateLine);
                string stringdata = Encoding.UTF8.GetString(byteDateLine, 0, recv);
                string ip = client.RemoteEndPoint.ToString();
                //获取客户端的IP和端口

                if (stringdata == "STOP")
                {
                    //当客户端终止连接时
                    showClientMsg(ip + "   " + now.ToString("G") + "  " + "已从服务器断开" + "\r\n");
                    break;
                }
                else if (stringdata.Trim() == "init")
                {
                    PPTController.Init();
                }
                else if (stringdata.Trim() == "first")
                {
                    PPTController.First();
                }
                else if (stringdata.Trim() == "last")
                {
                    PPTController.Last();
                }
                else if (stringdata.Trim() == "next")
                {
                    PPTController.Next();
                }
                else if (stringdata.Trim() == "prev")
                {
                    PPTController.Prev();
                }
                //显示客户端发送过来的信息
                showClientMsg(ip + "  " + now.ToString("G") + "   " + stringdata + "\r\n");
            }

        }

        //用来往msgTextBox框中显示消息
        public void showClientMsg(string msg)
        {
            //在线程里以安全方式调用控件
            if (msgTextBox.InvokeRequired)
            {
                MyInvoke _myinvoke = new MyInvoke(showClientMsg);
                msgTextBox.Invoke(_myinvoke, new object[] { msg });
            }
            else
            {
                msgTextBox.AppendText(msg);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //新建一个委托线程
            ThreadStart myThreadDelegate = new ThreadStart(Listen);
            //实例化新线程
            myThread = new Thread(myThreadDelegate);

            if (servicesStatus)
            {

                myThread.Start();
                statuBar.Text = "服务已启动，等待客户端连接";
                servicesStatus = false;
                startButton.Text = "停止服务";
            }
            else
            {
                //停止服务（绑定的套接字没有关闭,因此客户端还是可以连接上来）
                myThread.Interrupt();
                myThread.Abort();

                //showClientMsg("服务器已停止服务"+"\r\n");
                servicesStatus = true;
                startButton.Text = "开始服务";
                statuBar.Text = "服务已停止";

            }
        }
    }
}
