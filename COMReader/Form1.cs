using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;

namespace COMReader
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        private static Dictionary<string, string> dictionary;
        private Dictionary<string, string> commands;
        private BackgroundWorker workThread;
        private SerialPort port;

        private const int KEYEVENTF_EXTENDEDKEY = 0x0001; //Key down flag
        private const int KEYEVENTF_KEYUP = 0x0002; //Key up flag
        private const int VK_RCONTROL = 0xA3; //Right Control key code

        private bool receive = true;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            outputLabel.Text = "Avviamento sistema...";

            workThread = new BackgroundWorker();
            workThread.DoWork += new DoWorkEventHandler(serverWork);

            string[] lines = File.ReadAllLines(Path.Combine(Environment.CurrentDirectory, "settings.ini"));
            dictionary = new Dictionary<string, string>();
            commands = new Dictionary<string, string>();

            outputBox.AppendText(format(OutLevel.INFO, "File 'settings.ini' caricato\n"));

            bool apply = false;
            foreach (var str in lines) {
                str.Trim();

                if(str.StartsWith(";#"))
                {
                    apply = true;
                }

                if(str.StartsWith(";") || str.Equals(""))
                {
                    continue;
                }

                string[] property = str.Split('=');

                if(property.Length != 2)
                {
                    outputBox.AppendText(format(OutLevel.SEVERE, string.Format("La proprietà '{0}' non puo' essere caricata.\n", str)));
                    continue;
                }

                if(apply)
                {
                    commands.Add(property[0], property[1]);
                    outputBox.AppendText(format(OutLevel.INFO, string.Format("Comando '{0}' caricato con valore '{1}'\n", property[0], property[1])));
                } else
                {
                    dictionary.Add(property[0], property[1]);
                    outputBox.AppendText(format(OutLevel.INFO, string.Format("Proprietà '{0}' caricata con valore '{1}'\n", property[0], property[1])));
                }
            }

            workThread.RunWorkerAsync();
        }


        private void serverWork(object sender, DoWorkEventArgs e)
        {
            string propertyPort = dictionary["port"];

            port = new SerialPort()
            {
                PortName = propertyPort,
                ReceivedBytesThreshold = int.Parse(dictionary["bytes-threshold"]),
                Parity = Parity.None,
                StopBits = StopBits.One,
                BaudRate = int.Parse(dictionary["baud-rate"]),
                DtrEnable = false,
                RtsEnable = false,
                DataBits = 8
            };
            port.DataReceived += new SerialDataReceivedEventHandler(dataReceive);

            try
            {
                port.Open();
            }
            catch(Exception exception)
            {
                appendText("[Errore]:" + exception.Message);
                return;
            }

            setText("Sistema avviato, in ascolto...", outputLabel);
            appendText(format(OutLevel.INFO, string.Format("[Thread/{0}] Server avviato ed in ascolto sulla porta '{1}'\n", Thread.CurrentThread.ManagedThreadId, propertyPort)));
        }

        private char[] delims = new[] { '\r', '\n' };
        private void dataReceive(object sender, SerialDataReceivedEventArgs e)
        {
            if (!receive) return;

            SerialPort port = (SerialPort) sender;
            while (port.BytesToRead > 0)
            {
                byte[] buffer = new byte[port.BytesToRead];
                int count = port.Read(buffer, 0, port.BytesToRead);
                string cmdr = "";

                for(int i = 0; i < count; i++)
                {
                    cmdr += Convert.ToChar(buffer[i]);
                }

                string[] a = cmdr.Split(delims, StringSplitOptions.RemoveEmptyEntries);
                foreach(var k in a)
                {
                    appendText(format(OutLevel.INFO, "Comando ricevuto: " + k + "\n"));
                    k.Trim();
                    if (k == null || k == "") return;

                    performAction(k);
                }
            }
        }

        private void performAction(string command)
        {
            if(command.Length == 0)
            {
                return;
            }

            command.Trim();

            switch(command)
            {
                default: break;
                case "X1":
                    pressKey(commands["X1"], 1);
                    return;
                case "X10":
                    pressKey(commands["X10"], 1);
                    return;
                case "X100":
                    pressKey(commands["X100"], 1);
                    return;
            }

            string subcmd = command.Substring(0, 2);
            int amount = 0;
            
            try
            {
                amount = int.Parse(command.Substring(2));
            }
            catch(FormatException e)
            {
                appendText("[Errore]: " + e.Message);
            }

            if(!commands.ContainsKey(subcmd))
            {
                return;
            }

            pressKey(commands[subcmd], amount);
        }

        private void pressKey(string key, int amount)
        {
            var processes = Process.GetProcessesByName(dictionary["process-name"]);
            if (processes.Length == 0)
            {
                appendText("No Process found");
                return;
            }

            AutomationElement element = AutomationElement.FromHandle(processes[0].MainWindowHandle);
            if (element == null) return;
            element.SetFocus();

            byte value = byte.Parse(key, System.Globalization.NumberStyles.HexNumber);

            for(int i = 0; i < amount; i++)
            {
                keybd_event(value, 0, KEYEVENTF_EXTENDEDKEY, 0);
                keybd_event(value, 0, KEYEVENTF_KEYUP, 0);

                string mout = i + ", " + key + " , " + value.ToString() + " , " + (value >> 8).ToString();
                setText(mout, label3);
            }
        }

        private string format(OutLevel level, string s) {
            string result = "";

            switch(level)
            {
                default:
                case OutLevel.INFO:
                    result += "[INFO]: ";
                    break;
                case OutLevel.WARNING:
                    result += "[WARN]: ";
                    break;
                case OutLevel.SEVERE:
                    result += "[SEVERE]: ";
                    break;
            }

            result += s;

            return result;
        }

        private enum OutLevel { INFO, WARNING, SEVERE }

        delegate void appendTextCallback(string text);
        delegate void setTextCallback(string text, Label label);

        private void appendText(string text)
        {
            if(outputBox.InvokeRequired)
            {
                appendTextCallback callback = new appendTextCallback(appendText);
                Invoke(callback, new object[] { text });
            } else
            {
                outputBox.AppendText(text);
            }
        }

        public void setText(string text, Label label)
        {
            if (outputLabel.InvokeRequired)
            {
                setTextCallback callback = new setTextCallback(setText);
                Invoke(callback, new object[] { text, label });
            }
            else
            {
                label.Text = text;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            port.Close();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            outputBox.Clear();
        }

        private void OutputBox_TextChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked)
            {
                outputBox.SelectionStart = outputBox.Text.Length;
                outputBox.ScrollToCaret();
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            receive = !receive;

            if(receive)
            {
                button2.Text = "Ferma la ricezione dati";
                outputBox.AppendText(format(OutLevel.INFO, "Ricezione dati avviata\n"));
                port.Open();
            } else
            {
                button2.Text = "Abilita la ricezione dati";
                outputBox.AppendText(format(OutLevel.INFO, "Ricezione dati fermata\n"));
                port.Close();
            }
        }
    }
}
