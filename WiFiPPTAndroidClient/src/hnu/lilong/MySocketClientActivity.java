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
//    		DataInputStream dInputStream = new DataInputStream(socket.getInputStream());
//    		String msg = dInputStream.readUTF();
//    		
//    		EditText et = (EditText)findViewById(R.id.editText1);
//    		et.setText(msg);
    	} catch (Exception e) {
    		e.printStackTrace();
		}
    }
}