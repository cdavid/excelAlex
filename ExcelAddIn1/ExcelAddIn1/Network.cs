using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Sockets;
using System.IO;
using System.Windows.Forms;
using System.Net;
using System.Threading;

namespace ExcelAddIn1
{
    public class Network
    {
        private Socket s;
        private StreamReader sr;
        private StreamWriter sw;
        private NetworkStream stream;
        private Alex alx;


        public Network(Alex a)
        {
            this.alx = a;
            Thread conn = new Thread(new ThreadStart(connect));
            conn.Start();
        }

        public void stop()
        {
            this.sr.Close();
            this.sw.Close();
            this.stream.Close();
            this.s.Close();
        }

        public void send(String str)
        {
            try
            {
                if (stream.CanWrite)
                    sw.WriteLine(str);
            }
            catch (IOException ex)
            {
                //not quite sure if stuff might go wrong here                
            }
        }

        private void connect()
        {
            s = ConnectSocket(Dns.GetHostName(), 54117);
            if (s == null)
            {
                MessageBox.Show("Socket connect failed!!!");
                return;
            }
            stream = new NetworkStream(s);
            sr = new StreamReader(stream);
            sw = new StreamWriter(stream);
            sw.AutoFlush = true;

            Thread inRead = new Thread(new ThreadStart(read));
            inRead.Start();
        }

        private void read()
        {
            string response = "";
            try
            {
                while ((response = sr.ReadLine()) != null)
                {
                    this.alx.parseMessage(response);
                }
            }
            catch (IOException ex)
            {
                //we will reach this point if the socket closes while the ReadLine method is blocking
                //nothing to see here, move along
            }

        }

        private static Socket ConnectSocket(string server, int port)
        {
            Socket s = null;
            IPHostEntry hostEntry = null;

            // Get host related information.
            hostEntry = Dns.GetHostEntry(server);

            // Loop through the AddressList to obtain the supported AddressFamily. This is to avoid
            // an exception that occurs when the host IP Address is not compatible with the address family
            // (typical in the IPv6 case).
            foreach (IPAddress address in hostEntry.AddressList)
            {
                IPEndPoint ipe = new IPEndPoint(address, port);
                Socket tempSocket =
                    new Socket(ipe.AddressFamily, SocketType.Stream, ProtocolType.Tcp);

                tempSocket.Connect(ipe);

                if (tempSocket.Connected)
                {
                    s = tempSocket;
                    break;
                }
                else
                {
                    continue;
                }
            }
            return s;
        }
    }
}
