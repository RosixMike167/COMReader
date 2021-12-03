using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
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
        private SerialPort port;

        private const int KEYEVENTF_EXTENDEDKEY = 0x0001; //Key down flag
        private const int KEYEVENTF_KEYUP = 0x0002; //Key up flag

        private bool receive = true;
        private int trycount = 0;

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Show(object sender, EventArgs e)
        {
            outputBox.SelectionStart = outputBox.TextLength;
            outputBox.ScrollToCaret();
            outputBox.Invalidate();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            dictionary = new Dictionary<string, string>();
            commands = new Dictionary<string, string>();

            loadCOMList();
            loadSettings();
            createCOM(dictionary["port"]);
            openCOM("");
        }

        private bool isNumber(string v)
        {
            for (int i = 0; i < v.Length; i++) if ((v.ElementAt(i) < '0' || v.ElementAt(i) > '9') && (v.ElementAt(i)!='-' || v.ElementAt(i) != '+')) return false;
            return true;
        }

        private void loadCOMList()
        {
            listBox1.Items.Clear();
            listBox1.Items.AddRange(SerialPort.GetPortNames());
        }

        private void createCOM(string portCom)
        {
            if (dictionary.Count == 0 || commands.Count == 0)
            {
                outbox(format(OutLevel.WARNING, "Proprietà e Comandi non caricati...\n"),Color.Brown);
                return;
            }

            if (listBox1.Items.Count == 0)
            {
                outbox(format(OutLevel.WARNING, "Nessuna porta COM disponibile"),Color.Brown);
                return;
            }

            int brate = 115200, btrsh = 10;

            if (isNumber(dictionary["baud-rate"])) brate = int.Parse(dictionary["baud-rate"]);
            if (brate == 0) brate = 115200;
            if (isNumber(dictionary["bytes-threshold"])) btrsh = int.Parse(dictionary["bytes-threshold"]);
            if (btrsh < 2) btrsh = 10;

            port = new SerialPort()
            {
                PortName = portCom,
                ReceivedBytesThreshold = btrsh,
                Parity = Parity.None,
                StopBits = StopBits.One,
                BaudRate = brate,
                DtrEnable = false,
                RtsEnable = false,
                DataBits = 8
            };
            port.DataReceived += new SerialDataReceivedEventHandler(dataReceive);
        }

        private void closeCOM()
        {
            if (port == null) return;
            if (!port.IsOpen) return;

            try
            {
                port.Close();
                button2.Enabled = true;
                pictureBox1.Image = Properties.Resources.off;
            }
            catch (Exception exception)
            {
                outbox(format(OutLevel.ERROR, exception.Message + '\n'),Color.Red);
                return;
            }
        }

        private void openCOM(string portName)
        {
            if (port == null) return;
            if (port.IsOpen) closeCOM();

            if(portName != "") port.PortName = portName;

            try
            {
                port.Open();
                button2.Enabled = true;
                trycount = 0;
                outbox(format(OutLevel.INFO, string.Format("Server avviato ed in ascolto sulla porta '{0}'\n", port.PortName)),Color.Green);
                pictureBox1.Image = Properties.Resources.on;
            }
            catch (Exception exception)
            {
                button2.BackColor = Color.Red;
                button2.Enabled = false;
                trycount++;
                outbox(format(OutLevel.SEVERE, string.Format("Impossibile avviare il server sulla porta '{0}'\n", port.PortName)),Color.Brown);
                pictureBox1.Image = Properties.Resources.off;
                return;
            }

            button2.BackColor = Color.Green;
            return;
        }

        delegate void SetTextCallback(string text);
        private void toplabel(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.outputBox.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(toplabel);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.label1.Text=text;
            }
        }

        Color myColor=Color.Black;
        private void outbox(string text, Color c)
        {
            myColor = c;
            outbox(text);
        }
        private void outbox(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.outputBox.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(outbox);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                outputBox.SelectionColor = myColor;
                this.outputBox.AppendText(text);
            }
        }

        private char[] delims = new[] { '\r', '\n' };

        String DataInput = "";
        private void dataReceive(object sender, SerialDataReceivedEventArgs e)
        {
            if (!receive || !this.port.IsOpen) return;

            pictureBox2.Image = Properties.Resources.on;
            pictureBox2.Invalidate();

            SerialPort port = (SerialPort) sender;
            while (port.BytesToRead > 0)
            {
                DataInput += port.ReadExisting();

                string[] a = DataInput.Split(delims, StringSplitOptions.RemoveEmptyEntries);

                if (a.Length > 0)
                {
                    if (DataInput.LastIndexOf("\r\n") > 0) DataInput = DataInput.Substring(DataInput.LastIndexOf("\r\n")); else DataInput = "";
                    foreach (var k in a)
                    {
                        string k2 = k.Trim();
                        if (k2 != "") performAction(k2);
                    }
                }

            }
        }

        int weelpos = 0;
        private void performAction(string command)
        {
            command = command.Trim();

            pictureBox2.Image = Properties.Resources.off;
            pictureBox2.Invalidate();

            string subcmd = command.Trim();
            int amount = 1;
            
            if (command.Length < 2) return;
            if(subcmd.ElementAt(1)=='_')
            {
                toplabel(subcmd.Substring(2)+ " " + weelpos);
            }
            if(subcmd.Contains("OFF+") && subcmd.Length>4)
            {
                string a = command.Substring(command.IndexOf("/") + 1);
                
                if(a!="")
                {
                    if (isNumber(a)) {
                        try
                        {
                            weelpos = Convert.ToInt32(a);
                        } catch ( Exception exception) { weelpos = 0; }  
                        toplabel("OFF " + weelpos.ToString());
                        return;

                    }
                }
            }
            if (command.Substring(1, 1) == "+" || command.Substring(1, 1) == "-")
            {
                subcmd = command.Substring(0, 2);
                amount = 0;

                try
                {
                    bool a = false;
                    if (command.IndexOf("/")>0)
                    {
                        weelpos = Convert.ToInt32(command.Substring(command.IndexOf("/")+1));
                        command = command.Substring(0, command.IndexOf("/"));
                        a = true;
                    }
                    
                    amount = int.Parse(command.Substring(2));
                    
                    if (a)
                    {
                        weelpos *= amount;
                        toplabel(subcmd.Substring(0,1)+" "+weelpos.ToString());
                    }

                }
                catch (FormatException exception)
                {
                    if (checkBox2.Checked) outbox(format(OutLevel.ERROR, exception.Message + '\n'),Color.Brown);
                }
            }

            if (commands.ContainsKey(subcmd) && amount > 0)
            {
                pressKey(commands[subcmd], amount);

                if (checkBox2.Checked) outbox(format(OutLevel.INFO, "Comando: " + command+ '\n'),Color.BlueViolet);

                return;
            }

        }

        private void pressKey(string key, int amount)
        {
            if (key.Length == 1) key = "0" + key;

            string pname = dictionary["process-name"];

            if (pname != "")
            {
                var processes = Process.GetProcessesByName(pname);
                if (processes.Length == 0)
                {
                    outbox(format(OutLevel.ERROR, "Process " + pname + " not found\n"),Color.Brown);
                    return;
                }

                AutomationElement element = AutomationElement.FromHandle(processes[0].MainWindowHandle);
                if (element == null)
                {
                    outbox(format(OutLevel.ERROR, "Process " + pname + " hanldle Error"), Color.Brown);
                    return;
                }

                try {
                    element.SetFocus();
                }
                catch (Exception ex)
                {
                    outbox(format(OutLevel.ERROR, ex.Message + '\n'), Color.Red);
                }

            }

            int value;
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
            catch (FormatException exception)
            {
                outbox(format(OutLevel.ERROR, exception.Message + '\n'),Color.Red);
            }

            if (val1 == 0 && val2 == 0) return;
            for (int i = 0; i < amount; i++)
            {
                keybd_event(val1, val2, KEYEVENTF_EXTENDEDKEY, 0);
                keybd_event(val1, val2, KEYEVENTF_KEYUP, 0);
            }
        }

        private string format(OutLevel level, string s)
        {
            string result = "", datenow = DateTime.Now.ToString("HH:mm:ss");

            switch (level)
            {
                default:
                case OutLevel.INFO:
                    result += "[" + datenow + "] [INFO]: ";
                    break;
                case OutLevel.WARNING:
                    result += "[" + datenow + "] [WARN]: ";
                    break;
                case OutLevel.SEVERE:
                    result += "[" + datenow + "] [SEVERE]: ";
                    break;
                case OutLevel.ERROR:
                    result += "[" + datenow + "] [ERROR]: ";
                    break;
            }

            result += s;

            return result;
        }

        private enum OutLevel { INFO, WARNING, SEVERE, ERROR }

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
            if (checkBox1.Checked )
            {
                outputBox.SelectionStart = outputBox.Text.Length;
                outputBox.ScrollToCaret();
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Process.Start(Path.Combine(Environment.CurrentDirectory, "settings.ini"));
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            if (port == null) return;       // se non ci sono com disponibili port è null !!!
            if (port.IsOpen || !receive) return;

            outbox(format(OutLevel.WARNING, string.Format("Porta COM non disponibile, tentativo numero {0}...\n", trycount)),Color.Brown);
            openCOM("");
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            closeCOM();
            loadCOMList();
            loadSettings();
            createCOM(dictionary["port"]);
            openCOM("");
        }

        private void loadSettings()
        {
            dictionary.Clear();
            commands.Clear();

            string[] lines = {};

            try
            {
                lines = File.ReadAllLines(Path.Combine(Environment.CurrentDirectory, "settings.ini"));
                outbox(format(OutLevel.INFO, "File 'settings.ini' caricato\n"),Color.Green);
            }
            catch (Exception exception)
            {
                outbox(format(OutLevel.ERROR, exception.Message + '\n'),Color.Red);
                return;
            }

            bool apply = false;
            int sumerr = 0, nprop = 0, ncommands = 0;
            foreach (var str in lines)
            {
                string trimstr = str.Trim();

                if (trimstr.StartsWith(";#"))
                {
                    apply = true;
                }

                if (trimstr.StartsWith(";") || trimstr.Equals(""))
                {
                    continue;
                }

                string[] property = trimstr.Split('=');

                if (property.Length != 2 || property[0] == "" || property[1] == "")
                {
                    outbox(format(OutLevel.SEVERE, string.Format("La proprietà '{0}' non puo' essere caricata.\n", trimstr)),Color.Brown);
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
            
            outbox(format(OutLevel.INFO, string.Format("Proprietà Caricate: {0}, Comandi: {1}, Errori: {2}\n", nprop, ncommands, sumerr)),Color.Green);
            
        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            closeCOM();
            openCOM(listBox1.Text);
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            receive = !receive;

            if (receive)
            {
                openCOM("");
                button2.Text = "Ferma la ricezione dati";
                outbox(format(OutLevel.INFO, "Ricezione dati avviata\n"),Color.Green);
            }
            else
            {
                closeCOM();
                outbox(format(OutLevel.INFO, "Ricezione dati fermata\n"),Color.Brown);
                button2.Text = "Abilita la ricezione dati";
            }
        }

    }
}
