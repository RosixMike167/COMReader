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

        private bool receive = true;

        public Form1()
        {
            NewMethod();
        }

        private void NewMethod()
        {

            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            outputLabel.Text = "Avviamento sistema...";

            workThread = new BackgroundWorker();

            dictionary = new Dictionary<string, string>();
            commands = new Dictionary<string, string>();
            string[] lines = { };

            try
            {
                lines = File.ReadAllLines(Path.Combine(Environment.CurrentDirectory, "settings.ini"));
                outputBox.AppendText(format(OutLevel.INFO, "File 'settings.ini' caricato\n"));
            }
            catch (Exception exception)
            {
                appendText("[Errore]:" + exception.Message + '\n');
                return;
            }

            workThread.DoWork += new DoWorkEventHandler(serverWork);

            bool apply = false;
            int sumerr = 0, nprop = 0, ncommands = 0;
            foreach (var str in lines)
            {
                str.Trim();

                if (str.StartsWith(";#"))
                {
                    apply = true;
                }

                if (str.StartsWith(";") || str.Equals(""))
                {
                    continue;
                }

                string[] property = str.Split('=');

                if (property.Length != 2)
                {
                    outputBox.AppendText(format(OutLevel.SEVERE, string.Format("La proprietà '{0}' non puo' essere caricata.\n", str)));
                    sumerr++;
                    continue;
                }

                if (property[0] == "" || property[1] == "")
                {
                    outputBox.AppendText(format(OutLevel.SEVERE, string.Format("La proprietà '{0}' non puo' essere caricata.\n", str)));
                    sumerr++;
                    continue;
                }

                if (apply)
                {
                    commands.Add(property[0], property[1]);
                    outputBox.AppendText(format(OutLevel.INFO, string.Format("Comando '{0}' caricato con valore '{1}'\n", property[0], property[1])));
                    ncommands++;
                }
                else
                {
                    dictionary.Add(property[0], property[1]);
                    outputBox.AppendText(format(OutLevel.INFO, string.Format("Proprietà '{0}' caricata con valore '{1}'\n", property[0], property[1])));
                    nprop++;
                }
            }

            outputBox.AppendText(format(OutLevel.INFO, string.Format("Proprietà Caricate '{0}', Comandi '{1}', Errori '{2}'\n", nprop, ncommands, sumerr)));
            workThread.RunWorkerAsync();
        }

        bool isNumber(String v)
        {
            for (int i = 0; i < v.Length; i++) if (v.ElementAt(i) < '0' || v.ElementAt(i) > '9') return false;
            return true;
        }

        private void serverWork(object sender, DoWorkEventArgs e)
        {
            string propertyPort = dictionary["port"];
            int brate = 115200, btrsh = 10;

            if (isNumber(dictionary["baud-rate"])) brate = int.Parse(dictionary["baud-rate"]);
            if (brate == 0) brate = 115200;
            if (isNumber(dictionary["bytes-threshold"])) btrsh = int.Parse(dictionary["bytes-threshold"]);
            if (btrsh < 2) btrsh = 10;

            port = new SerialPort()
            {
                PortName = propertyPort,
                ReceivedBytesThreshold = btrsh,
                Parity = Parity.None,
                StopBits = StopBits.One,
                BaudRate = brate,
                DtrEnable = false,
                RtsEnable = false,
                DataBits = 8
            };
            port.DataReceived += new SerialDataReceivedEventHandler(dataReceive);

            openCOM();

            setText("Sistema avviato, in ascolto...", outputLabel);
            appendText(format(OutLevel.INFO, string.Format("[Thread/{0}] Server avviato ed in ascolto sulla porta '{1}'\n", Thread.CurrentThread.ManagedThreadId, propertyPort)));
        }

        void closeCOM()
        {
            if (port == null) return;
            if (!port.IsOpen) return;
            try
            {
                port.Close();
            }
            catch (Exception exception)
            {
                appendText("[Errore]:" + exception.Message + '\n');
                return;
            }
        }

        void openCOM()
        {
            if (port == null) return;
            if (port.IsOpen) closeCOM();
            try
            {
                port.Open();
            }
            catch (Exception exception)
            {
                appendText("[Errore]:" + exception.Message + '\n');
                button2.BackColor = Color.Red;
                return;
            }
            button2.BackColor = Color.Green;
        }

        private char[] delims = new[] { '\r', '\n' };
        private void dataReceive(object sender, SerialDataReceivedEventArgs e)
        {
            if (!receive) return;

            SerialPort port = (SerialPort)sender;
            while (port.BytesToRead > 0)
            {
                byte[] buffer = new byte[port.BytesToRead];
                int count = port.Read(buffer, 0, port.BytesToRead);
                string cmdr = "";

                for (int i = 0; i < count; i++)
                {
                    cmdr += Convert.ToChar(buffer[i]);
                }

                string[] a = cmdr.Split(delims, StringSplitOptions.None);
                foreach (var k in a)
                {
                    String k2 = k.Replace("\0", "");
                    k2 = k2.Trim();
                    //appendText("Seriale: " + k2 + ", len="+k2.Length+ "\n");
                    if (k2 != "") performAction(k2);
                }

            }
        }

        private void performAction(string command)
        {
            command = command.Trim();

            string subcmd = command.Trim();
            int amount = 1;

            if (command.Substring(1, 1) == "+" || command.Substring(1, 1) == "-")
            {
                subcmd = command.Substring(0, 2);
                amount = 0;

                try
                {
                    amount = int.Parse(command.Substring(2));
                }
                catch (FormatException e)
                {
                    appendText("[Errore]: " + e.Message + '\n');
                }
            }

            if (commands.ContainsKey(subcmd) && amount > 0)
            {
                pressKey(commands[subcmd], amount);
                appendText("\nComando: " + command);
                return;
            }

        }

        private void pressKey(string key, int amount)
        {
            if (key.Length == 1) key = "0" + key;
            String pname = dictionary["process-name"];
            if (pname != "")
            {
                var processes = Process.GetProcessesByName(pname);
                if (processes.Length == 0)
                {
                    appendText("Process " + pname + " not found\n");
                    return;
                }

                AutomationElement element = AutomationElement.FromHandle(processes[0].MainWindowHandle);
                if (element == null)
                {
                    appendText("Process " + pname + " hanldle Error");
                    return;
                }
                try { element.SetFocus(); } catch (Exception ex) { appendText(ex.Message + '\n'); }

            }

            int value = 0;
            byte val1 = 0, val2 = 0;
            try
            {
                if (key.Length > 3)
                {
                    value = int.Parse(key.Substring(0, 4), System.Globalization.NumberStyles.HexNumber);
                    val1 = (byte)value;
                    val2 = (byte)(val1 >> 8);
                }
                else
                {
                    val1 = byte.Parse(key.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                    val2 = 0;
                    value = val1;
                }
            }
            catch (FormatException e)
            {
                appendText("[Errore]: " + e.Message + '\n');
            }
            if (val1 == 0 && val2 == 0) return;
            for (int i = 0; i < amount; i++)
            {
                keybd_event(val1, val2, KEYEVENTF_EXTENDEDKEY, 0);
                keybd_event(val1, val2, KEYEVENTF_KEYUP, 0);

                string mout = i + ", " + key + " , " + value.ToString() + " , " + val1.ToString() + " , " + val2.ToString();
                setText(mout, label3);
            }
        }

        private string format(OutLevel level, string s)
        {
            string result = "";

            switch (level)
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
            if (outputBox.InvokeRequired)
            {
                appendTextCallback callback = new appendTextCallback(appendText);
                Invoke(callback, new object[] { text });
            }
            else
            {
                if (outputBox.Lines.Count() > 1000 ) outputBox.Clear();         
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
            if (port != null) closeCOM();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            outputBox.Clear();
        }

        private void OutputBox_TextChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked && checkBox1.Visible && checkBox1.Enabled)
            {
                outputBox.SelectionStart = outputBox.Text.Length;
                outputBox.ScrollToCaret();
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            receive = !receive;

            if (receive)
            {
                openCOM();
                button2.Text = "Ferma la ricezione dati";
                outputBox.AppendText(format(OutLevel.INFO, "Ricezione dati avviata\n"));
            }
            else
            {
                closeCOM();
                outputBox.AppendText(format(OutLevel.INFO, "Ricezione dati fermata\n"));
                button2.Text = "Abilita la ricezione dati";
            }
        }
    }
}
