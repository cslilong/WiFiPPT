WiFiPPT
=======

This project is using devices control the PPT slide presentation.

WiFiPPT 手机WiFi控制PPT播放
======================
###一、预期效果
PPT放映效果<br/>

![image](https://github.com/cslilong/WiFiPPT/raw/master/ppt.gif "WiFiPPT")

![image](https://github.com/cslilong/WiFiPPT/raw/master/WiFiPPTClient.png "WiFiPPTAndroidClient")

###二、C#控制PPT放映

1.既然要实现的程序是遥控幻灯片，这样我们就需要先获得幻灯片应用程序的，在PowerPoint对象模型中，Microsoft.Office.Interop.PowerPoint.Application代表Powerpoint应用程序，这点和Word、Excel和Outlook都是一样的。<br/>
2.获得了幻灯片应用程序对象之后，之后我们就需要获得幻灯片对象，因为我们遥控的是幻灯片，在PowerPoint对象模型中也提供了幻灯片对象，即Microsoft.Office.Interop.PowerPoint.Slide。由于幻灯片又是存在于演示文稿中的，所以我们要想获得幻灯片对象，就需要先获得演示文稿对象，Microsoft.Office.Interop.PowerPoint.Presentation 就是代表演示文稿对象。<br/>
3.获得幻灯片对象之后，我们就可以利用幻灯片对象的Select方法来进行幻灯片的切换,然而在阅读模式的情况下，不能用Select方法来进行翻页，此时需要另一种方式来实现，即调用 Microsoft.Office.Interop.PowerPoint.SlideShowView对象的First，Next,Last,Previous方法来进行幻灯片翻页。<br/>
``` C#
using System ;
using System .Collections. Generic;
using System .Linq;
using System .Text;
using PPt = Microsoft. Office.Interop .PowerPoint;
using System .Runtime. InteropServices;

namespace WiFiPPT
{
    public class PPTController
    {
        // 定义PowerPoint 应用程序对象
        static PPt .Application pptApplication;
        // 定义演示文稿对象
        static PPt .Presentation presentation;
        // 定义幻灯片集合对象
        static PPt .Slides slides;
        // 定义单个幻灯片对象
        static PPt .Slide slide;

        // 幻灯片的数量
        static int slidescount;
        // 幻灯片的索引
        static int slideIndex;

        // 获取ppt 的句柄
        public static bool Init()
        {
            // 必须先运行幻灯片，下面才能获得 PowerPoint应用程序，否则会出现异常
            // 获得正在运行的PowerPoint应用程序
            try
            {
                pptApplication = Marshal.GetActiveObject ("PowerPoint.Application") as PPt. Application;
            }
            catch
            {
                return false ;
            }
            if (pptApplication != null)
            {
                //获得演示文稿对象
                presentation = pptApplication .ActivePresentation;
                // 获得幻灯片对象集合
                slides = presentation .Slides;
                // 获得幻灯片的数量
                slidescount = slides .Count;
                // 获得当前选中的幻灯片
                try
                {
                    // 在普通视图下这种方式可以获得当前选中的幻灯片对象
                    // 然而在阅读模式下，这种方式会出现异常
                    slide = slides[pptApplication .ActiveWindow. Selection.SlideRange .SlideNumber];
                }
                catch
                {
                    // 在阅读模式下出现异常时，通过下面的方式来获得当前选中的幻灯片对象
                    slide = pptApplication.SlideShowWindows [1].View. Slide;
                }
            }
            return true ;
        }
        // First
        public static void First()
        {
            try
            {
                // 在普通视图中调用Select方法来选中第一张幻灯片
                slides[1].Select ();
                slide = slides [1];
            }
            catch
            {
                // 在阅读模式下使用下面的方式来切换到第一张幻灯片
                pptApplication.SlideShowWindows [1].View. First();
                slide = pptApplication .SlideShowWindows[1]. View.Slide ;
            }
        }
        // Last
        public static void Last()
        {
            try
            {
                slides[slidescount ].Select();
                slide = slides [slidescount];
            }
            catch
            {
                // 在阅读模式下使用下面的方式来切换到最后幻灯片
                pptApplication.SlideShowWindows [1].View. Last();
                slide = pptApplication .SlideShowWindows[1]. View.Slide ;
            }
        }
        // Next
        public static void Next()
        {
            slideIndex = slide .SlideIndex + 1;
            if (slideIndex > slidescount)
            {
                //MessageBox.Show("已经是最后一页了");
            }
            else
            {
                try
                {
                    slide = slides [slideIndex];
                    slides[slideIndex ].Select();
                }
                catch
                {
                    // 在阅读模式下使用下面的方式来切换到下一张幻灯片
                    pptApplication.SlideShowWindows [1].View. Next();
                    slide = pptApplication.SlideShowWindows [1].View. Slide;
                }
            }
        }
        // Prev
        public static void Prev()
        {
            slideIndex = slide .SlideIndex - 1;
            if (slideIndex >= 1)
            {
                try
                {
                    slide = slides [slideIndex];
                    slides[slideIndex ].Select();
                }
                catch
                {
                    // 在阅读模式下使用下面的方式来切换到上一张幻灯片
                    pptApplication.SlideShowWindows [1].View. Previous();
                    slide = pptApplication.SlideShowWindows [1].View. Slide;
                }
            }
            else
            {
                //MessageBox.Show("已经是第一页了");
            }
        }
    }
}
```
###三、C#的Socket通信

服务端多线程异步程序<br/>
``` C#
public bool servicesStatus = true;  // 开始停止服务按钮状态
public Thread myThread;       //声明一个线程实例
public Socket newsock;        //声明一个Socket 实例
public Socket server;
public Socket client;
public EndPoint remote;
public IPEndPoint localEP;
public delegate void MyInvoke(string str);

//监听函数
public void Listen()
{
    //初始化SOCKET 实例
    newsock = new Socket( AddressFamily.InterNetwork , SocketType. Stream, ProtocolType.Tcp );
    //允许SOCKET 被绑定在已使用的地址上。
    newsock.SetSocketOption (SocketOptionLevel. Socket, SocketOptionName.ReuseAddress , true);
    //初始化终结点实例
    localEP = new IPEndPoint( IPAddress.Any , int.Parse (portTextBox. Text.Trim ()));
    try
    {
        //绑定
        newsock.Bind (localEP);
        //监听
        newsock.Listen (10);
        //开始接受连接，异步。
        newsock.BeginAccept (new AsyncCallback(OnConnectRequest ), newsock);
    }
    catch (Exception ex)
    {
        //showClientMsg(ex.Message);
    }
}

//当有客户端连接时的处理
public void OnConnectRequest( IAsyncResult ar )
{
    //初始化一个SOCKET，用于其它客户端的连接
    server = (Socket )ar. AsyncState;
    client = server .EndAccept( ar);
    //将要发送给连接上来的客户端的提示字符串
    DateTimeOffset now = DateTimeOffset. Now;
    string strDateLine = "欢迎登录到服务器 ";
    Byte[] byteDateLine = System.Text .Encoding. UTF8.GetBytes (strDateLine);
    //将提示信息发送给客户端 ,并在服务端显示连接信息。
    remote = client .RemoteEndPoint;
    showClientMsg(client .RemoteEndPoint. ToString() + " 连接成功。" + now.ToString ("G") + "\r\n");
    client.Send (byteDateLine, byteDateLine.Length , 0);

    //等待新的客户端连接
    server.BeginAccept (new AsyncCallback(OnConnectRequest ), server);

    while (true )
    {
        int recv = client. Receive(byteDateLine );
        string stringdata = Encoding. UTF8.GetString (byteDateLine, 0, recv);
        string ip = client. RemoteEndPoint.ToString ();
        //获取客户端的IP和端口

        if (stringdata == "STOP")
        {
            //当客户端终止连接时
            showClientMsg(ip + "   " + now.ToString ("G") + "  " + " 已从服务器断开 " + "\r\n" );
            break;
        }
        else if (stringdata. Trim() == "init" )
        {
            PPTController.Init ();
        }
        else if (stringdata. Trim() == "first" )
        {
            PPTController.First ();
        }
        else if (stringdata. Trim() == "last" )
        {
            PPTController.Last ();
        }
        else if (stringdata. Trim() == "next" )
        {
            PPTController.Next ();
        }
        else if (stringdata. Trim() == "prev" )
        {
            PPTController.Prev ();
        }
        //显示客户端发送过来的信息
        showClientMsg(ip + "  " + now.ToString ("G") + "   " + stringdata + "\r\n");
    }

}

//用来往msgTextBox 框中显示消息
public void showClientMsg( string msg )
{
    //在线程里以安全方式调用控件
    if (msgTextBox .InvokeRequired)
    {
        MyInvoke _myinvoke = new MyInvoke(showClientMsg );
        msgTextBox.Invoke (_myinvoke, new object [] { msg });
    }
    else
    {
        msgTextBox.AppendText (msg);
    }
}

private void button1_Click( object sender , EventArgs e)
{
    //新建一个委托线程
    ThreadStart myThreadDelegate = new ThreadStart(Listen );
    //实例化新线程
    myThread = new Thread( myThreadDelegate);

    if (servicesStatus )
    {

        myThread.Start ();
        statuBar.Text = "服务已启动，等待客户端连接 ";
        servicesStatus = false ;
        startButton.Text = "停止服务 ";
    }
    else
    {
        //停止服务（绑定的套接字没有关闭 ,因此客户端还是可以连接上来）
        myThread.Interrupt ();
        myThread.Abort ();

        //showClientMsg("服务器已停止服务"+"\r\n");
        servicesStatus = true ;
        startButton.Text = "开始服务 ";
        statuBar.Text = "服务已停止 ";

    }
}
```
客户端android程序<br/>
``` Java
package hnu.lilong;

import java.io.BufferedWriter;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.net.Socket;

import android.app.Activity;
import android.content.Context;
import android.net.wifi.WifiInfo;
import android.net.wifi.WifiManager;
import android.os.Bundle;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Button;
import android.widget.EditText;

public class MySocketClientActivity extends Activity {
    /** Called when the activity is first created. */
     private Button sendBtn;
     private Button getipBtn;
     private Button initBtn;
     private Button firstBtn;
     private Button lastBtn;
     private Button prevBtn;
     private Button nextBtn;
     private EditText et;
     private EditText ServerIpEt;
     private String serverIp;
     private Socket socket;
     private PrintWriter out = null;
    
    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.main);
        //connectToServer();
       
        et = (EditText)findViewById(R.id.editText1);
        ServerIpEt = (EditText)findViewById(R.id.ServeripEditText);
        sendBtn = (Button)findViewById(R.id.send);
        sendBtn.setOnClickListener(new Sendlistener());
       
        getipBtn = (Button)findViewById(R.id.getip);
        getipBtn.setOnClickListener(new Sendlistener1());
       
        initBtn = (Button)findViewById(R.id.init);
        initBtn.setOnClickListener(new InitSendlistener());
       
        firstBtn = (Button)findViewById(R.id.first);
        firstBtn.setOnClickListener(new FirstSendlistener());
       
        lastBtn = (Button)findViewById(R.id.last);
        lastBtn.setOnClickListener(new LastSendlistener());
       
        prevBtn = (Button)findViewById(R.id.prev);
        prevBtn.setOnClickListener(new PrevSendlistener());
       
        nextBtn = (Button)findViewById(R.id.next);
        nextBtn.setOnClickListener(new Nextendlistener());
    }
   
    private String intToIp(int i) {    
        return (i & 0xFF ) + "." +
         ((i >> 8 ) & 0xFF) + "." +
         ((i >> 16 ) & 0xFF) + "." +
         ( i >> 24 & 0xFF) ;
     }
   
    private String getServerIp(int i) {
         return (i & 0xFF ) + "." +
         ((i >> 8 ) & 0xFF) + "." +
         ((i >> 16 ) & 0xFF) + ".1";
    }
   
    class Sendlistener1 implements OnClickListener{
        public void onClick(View v) {
             //获取wifi服务
             WifiManager wifiManager = (WifiManager) getSystemService(Context.WIFI_SERVICE);
            //判断wifi是否开启
            if (!wifiManager.isWifiEnabled()) {
                 wifiManager.setWifiEnabled(true);
            }
            WifiInfo wifiInfo = wifiManager.getConnectionInfo();
            int ipAddress = wifiInfo.getIpAddress();
            String ip = intToIp(ipAddress);
            et.setText(ip);
            serverIp = getServerIp(ipAddress);
            ServerIpEt.setText(serverIp);
            connectToServer();
            getipBtn.setEnabled(false);
        }
    }
   
    class Sendlistener implements OnClickListener{ 
        public void onClick(View v) {
             try {
                  Send("shutdown");
               } catch (Exception e) {
                    e.printStackTrace();
               }
        }
       
    }
   
    class InitSendlistener implements OnClickListener{ 
        public void onClick(View v) {
             try {
                  Send("init");
               } catch (Exception e) {
                    e.printStackTrace();
               }
        }
    }
   
    class FirstSendlistener implements OnClickListener{ 
        public void onClick(View v) {
             try {
                  Send("first");
               } catch (Exception e) {
                    e.printStackTrace();
               }
        }
    }
   
    class LastSendlistener implements OnClickListener{ 
        public void onClick(View v) {
             try {
                  Send("last");
               } catch (Exception e) {
                    e.printStackTrace();
               }
        }
    }
    class PrevSendlistener implements OnClickListener{ 
        public void onClick(View v) {
             try {
                  Send("prev");
               } catch (Exception e) {
                    e.printStackTrace();
               }
        }
    }
    class Nextendlistener implements OnClickListener{ 
        public void onClick(View v) {
             try {
                  Send("next");
               } catch (Exception e) {
                    e.printStackTrace();
               }
        }
    }
   
    public void Send(String msg) {
         if (socket.isConnected()) {
            if (!socket.isOutputShutdown()) {
                out.println(msg);
            }
        }
    }
   
    public void connectToServer(){
         try {
              socket = new Socket(serverIp, 8888);
              out = new PrintWriter(new BufferedWriter(new OutputStreamWriter(socket.getOutputStream())), true);
//              DataInputStream dInputStream = new DataInputStream(socket.getInputStream());
//              String msg = dInputStream.readUTF();
//             
//              EditText et = (EditText)findViewById(R.id.editText1);
//              et.setText(msg);
         } catch (Exception e) {
              e.printStackTrace();
          }
    }
}
```
###项目文件说明：
服务端项目：WiFiPPT文件夹<br/>
android客户端文件：WiFiPPTAndroidClient文件夹<br/>
服务端可执行文件：WiFiPPT.exe（需要.net framework 2.0）<br/>
客户端Android的apk文件：WiFiPPTAndroidClient.apk<br/>


