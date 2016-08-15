using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Xml.Serialization;
using System.Configuration;
using System.Threading;
using System.Diagnostics;
using GemBox.Spreadsheet;
using System.IO;

namespace WindowsFormsApplication5
{
    public partial class Form1 : Form
    {
       // int numVal = 0;
        List<string> buttensofEnabledevices = new List<string>();
        List<string> IdsofEnabledevices = new List<string>(); 
        List<string> PortsofEnabledevices = new List<string>();
        List<string> CTsofEnabledevices = new List<string>();
        public List<string> AllconnectedPorts = new List<string>();
        public string firmwareName;
        public static string CallibrationDataFolderPath ="None";

        public int[] CTarray = { 30, 60, 100, 200, 400, 600, 1000, 1200, 1500, 3000 };
        public Form1()
        {
            InitializeComponent();
            this.MinimumSize = new Size(1400, 650);
            getportsClass vari = new getportsClass();  //get all connected ports in to array
            List<string> list1 = vari.GetAllPorts();

            //list one contains ports list
            //for (int i = 1; i <= 30; i++)
            //{
            
            for (int i = 2; i < 20; i++){
                string a=""+i;
                comboBox62.Items.Add(a);
            }
            comboBox62.SelectedIndex=2;
            
                foreach (var a in list1)
                {
                  //  string CombobxName;

                    comboBox31.Items.Add(a); comboBox32.Items.Add(a); comboBox40.Items.Add(a); comboBox33.Items.Add(a); comboBox34.Items.Add(a);
                    comboBox35.Items.Add(a); comboBox36.Items.Add(a); comboBox37.Items.Add(a); comboBox38.Items.Add(a); comboBox39.Items.Add(a);
                    comboBox41.Items.Add(a); comboBox42.Items.Add(a); comboBox43.Items.Add(a); comboBox44.Items.Add(a); comboBox45.Items.Add(a);
                    comboBox46.Items.Add(a); comboBox47.Items.Add(a); comboBox48.Items.Add(a); comboBox49.Items.Add(a); comboBox50.Items.Add(a);
                    comboBox51.Items.Add(a); comboBox52.Items.Add(a); comboBox53.Items.Add(a); comboBox54.Items.Add(a); comboBox55.Items.Add(a);
                    comboBox56.Items.Add(a); comboBox57.Items.Add(a); comboBox58.Items.Add(a); comboBox59.Items.Add(a); comboBox60.Items.Add(a);
                    comboBox61.Items.Add(a);

                }

                

        }
        

        private void Form1_Load(object sender, EventArgs e)
        {
            addCtsToComboBoxes(CTarray, comboBox1); addCtsToComboBoxes(CTarray, comboBox16);
            addCtsToComboBoxes(CTarray, comboBox2); addCtsToComboBoxes(CTarray, comboBox17);
            addCtsToComboBoxes(CTarray, comboBox3); addCtsToComboBoxes(CTarray, comboBox18);
            addCtsToComboBoxes(CTarray, comboBox4); addCtsToComboBoxes(CTarray, comboBox19);
            addCtsToComboBoxes(CTarray, comboBox5); addCtsToComboBoxes(CTarray, comboBox20);
            addCtsToComboBoxes(CTarray, comboBox6); addCtsToComboBoxes(CTarray, comboBox21);
            addCtsToComboBoxes(CTarray, comboBox7); addCtsToComboBoxes(CTarray, comboBox22);
            addCtsToComboBoxes(CTarray, comboBox8); addCtsToComboBoxes(CTarray, comboBox23);
            addCtsToComboBoxes(CTarray, comboBox9); addCtsToComboBoxes(CTarray, comboBox24);
            addCtsToComboBoxes(CTarray, comboBox10); addCtsToComboBoxes(CTarray, comboBox25);
            addCtsToComboBoxes(CTarray, comboBox11); addCtsToComboBoxes(CTarray, comboBox26);
            addCtsToComboBoxes(CTarray, comboBox12); addCtsToComboBoxes(CTarray, comboBox27);
            addCtsToComboBoxes(CTarray, comboBox13); addCtsToComboBoxes(CTarray, comboBox28);
            addCtsToComboBoxes(CTarray, comboBox14); addCtsToComboBoxes(CTarray, comboBox29);
            addCtsToComboBoxes(CTarray, comboBox15); addCtsToComboBoxes(CTarray, comboBox30);


            for (int i = 1; i <= 30; i++)
            {
                string retrievingName = "button" + i + "_backcolor";
                string retrievingcIdvalues = "textbox" + i + "_value";
                string retrievingPortCombobox = "combobox" + (i + 30) + "_port_value";
                string retrievingCTCombobox="combobox"+i+"_CT_value";
                string buttonName = "button" + i;
                string textBoxname = "textBox" + i;
                string comboboxPort = "comboBox" + (30 + i);
                string comboCT = "comboBox" + i;
                
                //get btton colors
                System.Drawing.Color myColor = System.Drawing.ColorTranslator.FromHtml(ConfigurationManager.AppSettings[retrievingName].ToString());
                this.Controls[buttonName].BackColor = myColor;
                
                //get textbox values
                this.Controls[textBoxname].Text = ConfigurationManager.AppSettings[retrievingcIdvalues].ToString();

                //get port values
                this.Controls[comboboxPort].Text = ConfigurationManager.AppSettings[retrievingPortCombobox].ToString();

                //get CT values
                this.Controls[comboCT].Text = ConfigurationManager.AppSettings[retrievingCTCombobox].ToString();
            }
            comboBox61.Text = ConfigurationManager.AppSettings["comboBox61_port_value"].ToString();
            



        }

        public IEnumerable<Control> GetAll(Control control, Type type)
        {
            var controls = control.Controls.Cast<Control>();

            return controls.SelectMany(ctrl => GetAll(ctrl, type))
                                      .Concat(controls)
                                      .Where(c => c.GetType() == type);
        }

       
        //Enable all
        private void button32_Click(object sender, EventArgs e)
        {
            var buttons = GetAll(this, typeof(Button));
            foreach (Control c in buttons)
                if(c.Text.Equals(""))
                    c.BackColor = Color.Green;
        }
        //Disable all
        private void button33_Click(object sender, EventArgs e)
        {
            var buttons = GetAll(this, typeof(Button));
            foreach (Control c in buttons)
                if (c.Text.Equals(""))
                    c.BackColor = Color.Red;

        }


        //event handler for all buttons when click on it 
        //handling multiple enable buttons at the same time 
        private void button1_Click(object sender, EventArgs e)
        {
            var btn = (Button)sender;
            string[] array = btn.Name.Split('n');
            int buttonNumber = Int32.Parse(array[1]);
            int comboBoxnumber = buttonNumber + 30;
            string ComboboxName = "comboBox" + comboBoxnumber;
            string ComboBoxContain = this.Controls[ComboboxName].Text;
            switch (btn.Name)
            {
                case "button1":

                    if (button1.BackColor == Color.Green)
                    {
                        button1.BackColor = Color.Red;

                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button1.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }

                    break;
                case "button2":
                    if (button2.BackColor == Color.Green)
                    {
                        button2.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button2.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button3":
                    if (button3.BackColor == Color.Green)
                    {
                        button3.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button3.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button4":
                    if (button4.BackColor == Color.Green)
                    {
                        button4.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");

                    }
                    else
                    {
                        button4.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button5":
                    if (button5.BackColor == Color.Green)
                    {
                        button5.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button5.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button6":
                    if (button6.BackColor == Color.Green)
                    {
                        button6.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button6.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button7":
                    if (button7.BackColor == Color.Green)
                    {
                        button7.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button7.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button8":
                    if (button8.BackColor == Color.Green)
                    {
                        button8.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button8.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button9":
                    if (button9.BackColor == Color.Green)
                    {
                        button9.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button9.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button10":
                    if (button10.BackColor == Color.Green)
                    {
                        button10.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button10.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button11":
                    if (button11.BackColor == Color.Green)
                    {
                        button11.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button11.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button12":
                    if (button12.BackColor == Color.Green)
                    {
                        button12.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button12.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button13":
                    if (button13.BackColor == Color.Green)
                    {
                        button13.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button13.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button14":
                    if (button14.BackColor == Color.Green)
                    {
                        button14.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button14.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button15":
                    if (button15.BackColor == Color.Green)
                    {
                        button15.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button15.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button16":
                    if (button16.BackColor == Color.Green)
                    {
                        button16.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button16.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button17":
                    if (button17.BackColor == Color.Green)
                    {
                        button17.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button17.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button18":
                    if (button18.BackColor == Color.Green)
                    {
                        button18.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button18.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button19":
                    if (button19.BackColor == Color.Green)
                    {
                        button19.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button19.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button20":
                    if (button20.BackColor == Color.Green)
                    {
                        button20.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button20.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button21":
                    if (button21.BackColor == Color.Green)
                    {
                        button21.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button21.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button22":
                    if (button22.BackColor == Color.Green)
                    {
                        button22.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button22.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button23":
                    if (button23.BackColor == Color.Green)
                    {
                        button23.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button23.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button24":
                    if (button24.BackColor == Color.Green)
                    {
                        button24.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button24.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button25":
                    if (button25.BackColor == Color.Green)
                    {
                        button25.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button25.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button26":
                    if (button26.BackColor == Color.Green)
                    {
                        button26.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button26.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button27":
                    if (button27.BackColor == Color.Green)
                    {
                        button27.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button27.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button28":
                    if (button28.BackColor == Color.Green)
                    {
                        button28.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button28.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button29":
                    if (button29.BackColor == Color.Green)
                    {
                        button29.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button29.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                case "button30":
                    if (button30.BackColor == Color.Green)
                    {
                        button30.BackColor = Color.Red;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_OFF");
                    }
                    else
                    {
                        button30.BackColor = Color.Green;
                        lightTheDeviceForRecognize(ComboBoxContain, "LED_ON");
                    }
                    break;
                default:
                    break;
            }
        }

        //add CT values to CT comboboxes
        public void addCtsToComboBoxes(Array array, ComboBox c)
        {
            foreach (var a in array)
            {
                c.Items.Add(a);
            }
        }

        //adding all selected COM ports to a list called "AllconnectedPorts"
        public void getPortslist() {
            string CombobxName;
            string portStatus;
            AllconnectedPorts.Clear();
            for (int i = 31; i <= 60; i++)
            {
                 CombobxName= "comboBox" + i;
                 portStatus = "button" + (i - 30);
                 if (this.Controls[portStatus].BackColor != Color.Red)
                 {
                     AllconnectedPorts.Add(this.Controls[CombobxName].Text);
                 }
            }

            
        }

        //auto fill ids + get enable devices Ids,Cts ,ports ,they are added to arrays in order
        private void button31_Click(object sender, EventArgs e)
        {
            
            int firstTetboxvalue = Int32.Parse(textBox1.Text);

            int global = -1;
            for (int i = 1; i <= 30; i++)
            {
                string TBName = "textBox" + (i);
                string buttonName = "button" + (i);
                string portName="comboBox"+(30+i);
                string CTName = "comboBox" + i;
                
                if (this.Controls[buttonName].BackColor == Color.Green)
                {
                    global=global+1;
                    this.Controls[TBName].Text = (firstTetboxvalue+global).ToString();
                  //  buttensofEnabledevices.Add(buttonName);
                  //  IdsofEnabledevices.Add(this.Controls[TBName].Text);
                  //  PortsofEnabledevices.Add(this.Controls[portName].Text);
                  //  CTsofEnabledevices.Add(this.Controls[CTName].Text);
                }
            }
           // MessageBox.Show(buttensofEnabledevices[1]);

        }

        //here we use Allconnectedports which user have assigned to devices(which computer gives you
        private void button34_Click(object sender, EventArgs e)
        {
            getPortslist();


            foreach (var por in AllconnectedPorts)
            {
                
                using (SerialPort port = new SerialPort(por))
                {
                    // Set the properties.
                    try
                    {
                        port.BaudRate = 300;
                        port.Parity = Parity.None;
                        port.ReadTimeout = 10;
                        port.StopBits = StopBits.One;

                        // Write a message into the port.
                        port.Open();
                        int i = 0;
                        while (i < 6)
                        {
                            // byte[] bytesToSend = new byte[1] {0x40};

                            port.Close();
                            port.Open();
                            port.Write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
                            port.Close();
                            Thread.Sleep(1000);
                            i++;
                        }

                        port.Close();
                        // Console.WriteLine("Wrote to the port.");
                        //Console.ReadLine();
                    }
                    catch { MessageBox.Show(port.PortName+" port is not conected"); }
                }
                 
            }
                

        }


        //from this function add states to the config file when save state clicks
        private void button35_Click(object sender, EventArgs e)
        {
            passControlsForsavestate();
        }



        public void passControlsForsavestate() {
             Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);
             for (int i = 1; i <= 30; i++) 
             {
                 
                 string buttonName = "button" + i;
                 string savingName = "button" + i + "_backcolor";
                 string Textvalue = "textBox" + i;
                 string textboxsavingName = "textBox" + i+"_value";
                 string comboPortName = "comboBox" + (30 + i);
                 string comboPortsavingName = "comboBox" + (30 + i) + "_port_value";
                 string comboCtName = "comboBox" + i;
                 string comboCTsavingName = "comboBox" + i + "_CT_value";

                 config.AppSettings.Settings[savingName].Value = this.Controls[buttonName].BackColor.Name.ToString();
                 config.AppSettings.Settings[textboxsavingName].Value = this.Controls[Textvalue].Text.ToString();
                 config.AppSettings.Settings[comboPortsavingName].Value = this.Controls[comboPortName].Text.ToString();
                 config.AppSettings.Settings[comboCTsavingName].Value = this.Controls[comboCtName].Text.ToString();
             }
             config.AppSettings.Settings["comboBox61_port_value"].Value = comboBox61.Text.ToString();
             
             config.AppSettings.Settings["comboBox62_port_value"].Value = comboBox62.Text.ToString();
             config.Save(ConfigurationSaveMode.Modified);
        }


        //this use to make the multi dimentional list of gatherd data
        public class EnableListDeviceInformation
        {
            public string DeviceId { get; set; }
            public string Ct { get; set; }
            public string Port { get; set; }
        }

        //if user dont use the autofill button ,here catch the enable buttons
        //execute button.. get enable devices data and store them in one multi row array .send in to the next process to another class
        private void button36_Click(object sender, EventArgs e)
        {   
            //for catch enabled buttons(only need when user not click auto fill)

          //  if (autofillclicked == false)
            {
                for (int i = 1; i <= 30; i++)
                {
                    string TBName = "textBox" + (i);
                    string buttonName = "button" + (i);
                    string portName = "comboBox" + (30 + i);
                    string CTName = "comboBox" + i;

                    if (this.Controls[buttonName].BackColor == Color.Green)
                    {
                        buttensofEnabledevices.Add(buttonName);
                        IdsofEnabledevices.Add(this.Controls[TBName].Text);
                        PortsofEnabledevices.Add(this.Controls[portName].Text);
                        CTsofEnabledevices.Add(this.Controls[CTName].Text);
                    }
                }
            }
            //here make the multi dimentional list so we can pass this to external class and 
            //do the process there
            var devicesData = new List<EnableListDeviceInformation>();
            
            for (int i = 0; i<IdsofEnabledevices.Count ; i++)
            {
                EnableListDeviceInformation temp = new EnableListDeviceInformation();
                temp.DeviceId = IdsofEnabledevices[i];
                temp.Ct = CTsofEnabledevices[i];
                temp.Port = PortsofEnabledevices[i];
                devicesData.Add(temp);
            }
            
            ExecutionOfProgramme clas = new ExecutionOfProgramme();
            clas.NumberOfAttemptsForData = 2;
            try
            {

                clas.NumberOfAttemptsForData = Int32.Parse(comboBox62.Text);
            }
            catch (Exception ex) { };
            //upload the first firmware (calibraion)
            clas.uploadFirmwares(devicesData, "Cal_FW_test10.ino.hex", label8);    
            
                var stopwatch = Stopwatch.StartNew();
                Thread.Sleep(6000);
                stopwatch.Stop();


            //write to the main controller port
            string mainControllerPort = comboBox61.Text;
            SerialPort port = new SerialPort(mainControllerPort);

             //Set the properties.
            port.BaudRate = 115200;
            port.Parity = Parity.Even;
            port.ReadTimeout = 10;
            port.StopBits = StopBits.One;
            port.Close();
            try
            {
                port.Open();
                //heaters,inducters,voltage
                port.Write("CON1,127,255,191,0,\n"); //
                Thread.Sleep(2000);
                port.Write("CON1,127,255,191,1,\n");    //first time for current
                port.Close();
            }
            catch (Exception ex) { this.excecutionErrorSummery(); System.Environment.Exit(1); }

            Thread.Sleep(2000);

            ////(b1+6*h+MC,I*6+B3+B2,V-transformmer,tiack)
            //write to the enabled ports that going to do the current calibraion.we can send the writing string later
            //  label11.Text = "abc";
            clas.DotheProcessAndReadvalues(devicesData, label8, label11, label12);//main controller changed inside the DotheProcess
            clas.doTheCurrentCalculations(devicesData, mainControllerPort, label11, label12, label8);
            try
            {
                port.Open();
                port.Write("CON1,159,255,191,1,\n");
                Thread.Sleep(1000);
                port.Write("CON1,159,255,191,0,\n");
                Thread.Sleep(1000);
                port.Write("CON1,127,255,254,0,\n");   //first time for voltage without main
                Thread.Sleep(1000);
                port.Write("CON1,127,255,254,1,\n");   //first time for voltage withmain
                Thread.Sleep(1000);
                port.Write("CON1,127,255,254,1,\n");
                port.Close();
            }
            catch (Exception ex) { this.excecutionErrorSummery(); System.Environment.Exit(1); }
            Thread.Sleep(4000);



            clas.VoltageDotheProcessAndReadvalues(devicesData, label8, label11, label12);
            clas.doTheVoltageCalculations(devicesData, mainControllerPort, label11, label12, label8);
            try
            {
                port.Open();

                port.Write("CON1,127,255,191,1,\n");
                Thread.Sleep(1000);
                port.Write("CON1,127,255,191,0,\n");
                Thread.Sleep(1000);
                port.Write("CON1,187,255,191,0,\n");
                Thread.Sleep(1000);//first time for phase without main
                port.Write("CON1,187,255,191,1,\n");
                //first time for phase with main
                port.Close();
            }
            catch (Exception ex) { this.excecutionErrorSummery(); System.Environment.Exit(1); }
            Thread.Sleep(8000);



            clas.PhaseDoTheProcessAndReadValues(devicesData, label8, label11, label12);
            clas.DothePhaseCalculations(devicesData, mainControllerPort, label11, label12, label8);
            ////commented
            try
            {
                port.Open();
                //change the main

                var sto3 = Stopwatch.StartNew();

                sto3.Stop();
                port.Close();
            }
            catch(Exception ex) { }

            clas.ActiveEnergyDoTheProcessAndReadValues(devicesData, mainControllerPort, label8, label11, label12);
            clas.DotheActiveEnergyCalculations(devicesData, mainControllerPort, label8, label11, label12);

            


            ReactiveEnergyExecution clas2 = new ReactiveEnergyExecution();

            clas2.ReactiveEnergyDoTheProcessAndReadValues(devicesData, mainControllerPort, label8, label11, label12);
            clas2.DotheReactiveEnergyCalculations(devicesData, mainControllerPort, label11, label12, label8);
            label11.Text = "";
            label11.Refresh();
            label12.Text = "";
            label12.Refresh();
            label8.Text = "Done!";
            label8.Refresh();
        }

        

      

        public void label8_Click(object sender, EventArgs e)
        {

        }
       
            public string label11CurrentClass
            {

                get
                {
                    return label11.Text;
                }
                set
                {
                    label11.Text = value;
                    label11.Refresh();
                }
            }

            public string label12VoltageClass
            {

                get
                {
                    return label12.Text;
                }
                set
                {
                    label12.Text = value;
                    label12.Refresh();
                }
            }

            public string label8class
            {

                get
                {
                    return label8.Text;
                }
                set
                {
                    label8.Text = value;
                    label8.Refresh();
                }
            }

            protected override void OnSizeChanged(EventArgs e)
            {
                base.OnSizeChanged(e);
                if (this.WindowState.Equals(FormWindowState.Minimized))
                {
                    this.Size = Screen.PrimaryScreen.WorkingArea.Size;
                }
                else if (this.WindowState.Equals(FormWindowState.Minimized))
                {
                    this.WindowState = FormWindowState.Minimized;
                }
            }

            private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
            {


            }

            private void button37_Click(object sender, EventArgs e)
            {
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                if (DialogResult.OK == fbd.ShowDialog())
                    Form1.CallibrationDataFolderPath = fbd.SelectedPath;
                else Form1.CallibrationDataFolderPath = "None";

            }

            private void button37_Click_1(object sender, EventArgs e)
            {
                this.addPortsToComboboxes();
            }
            private void addPortsToComboboxes()
            {
                comboBox31.Items.Clear(); comboBox32.Items.Clear(); comboBox33.Items.Clear(); comboBox34.Items.Clear(); comboBox35.Items.Clear();
                comboBox36.Items.Clear(); comboBox37.Items.Clear(); comboBox38.Items.Clear(); comboBox39.Items.Clear(); comboBox40.Items.Clear();
                comboBox61.Items.Clear();
                comboBox41.Items.Clear(); comboBox42.Items.Clear(); comboBox43.Items.Clear(); comboBox44.Items.Clear(); comboBox45.Items.Clear();
                comboBox46.Items.Clear(); comboBox47.Items.Clear(); comboBox48.Items.Clear(); comboBox49.Items.Clear(); comboBox50.Items.Clear();
                comboBox51.Items.Clear(); comboBox52.Items.Clear(); comboBox53.Items.Clear(); comboBox54.Items.Clear(); comboBox55.Items.Clear();
                comboBox56.Items.Clear(); comboBox57.Items.Clear(); comboBox58.Items.Clear(); comboBox59.Items.Clear(); comboBox60.Items.Clear();
                getportsClass vari = new getportsClass();  //get all connected ports in to array
                List<string> list1 = vari.GetAllPorts();

                //list one contains ports list
                //for (int i = 1; i <= 30; i++)
                //{
                foreach (var a in list1)
                {
                    //  string CombobxName;
                    

                    comboBox31.Items.Add(a); comboBox32.Items.Add(a); comboBox40.Items.Add(a); comboBox33.Items.Add(a); comboBox34.Items.Add(a);
                    comboBox35.Items.Add(a); comboBox36.Items.Add(a); comboBox37.Items.Add(a); comboBox38.Items.Add(a); comboBox39.Items.Add(a);
                    comboBox41.Items.Add(a); comboBox42.Items.Add(a); comboBox43.Items.Add(a); comboBox44.Items.Add(a); comboBox45.Items.Add(a);
                    comboBox46.Items.Add(a); comboBox47.Items.Add(a); comboBox48.Items.Add(a); comboBox49.Items.Add(a); comboBox50.Items.Add(a);
                    comboBox51.Items.Add(a); comboBox52.Items.Add(a); comboBox53.Items.Add(a); comboBox54.Items.Add(a); comboBox55.Items.Add(a);
                    comboBox56.Items.Add(a); comboBox57.Items.Add(a); comboBox58.Items.Add(a); comboBox59.Items.Add(a); comboBox60.Items.Add(a);
                    comboBox61.Items.Add(a);

                }
            }

            private void button38_Click(object sender, EventArgs e)
            {

                this.uploardFirmwares();

            }
            private void uploardFirmwares()
            {
                getportsClass vari = new getportsClass();  //get all connected ports in to array
                List<string> list1 = vari.GetAllPorts();
                
                //here make the multi dimentional list so we can pass this to external class and 
                //do the process there
                var devicesData = new List<EnableListDeviceInformation>();
                string mainController = this.comboBox61.Text;
                list1.Remove(mainController);
                if (MessageBox.Show("Have you selected the main contoller?", "Warning!", MessageBoxButtons.YesNo, MessageBoxIcon.Question)== DialogResult.Yes)
                {
                    
                        for (int i = 0; i < list1.Count; i++)
                        {
                            EnableListDeviceInformation temp = new EnableListDeviceInformation();
                            temp.Port = list1[i];
                            devicesData.Add(temp);
                        }

                        ExecutionOfProgramme clas = new ExecutionOfProgramme();
                        //upload the first firmware (calibraion)
                        clas.uploadFirmwares(devicesData, "Cal_FW_test10.ino.hex", label8);

                        var stopwatch = Stopwatch.StartNew();
                        Thread.Sleep(6000);
                        stopwatch.Stop();
                    }
                
            }

            private void button39_Click(object sender, EventArgs e)
            {
                List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> MainControllerList = new List<EnableListDeviceInformation>();
                EnableListDeviceInformation mainController = new EnableListDeviceInformation();
                mainController.Port = this.comboBox61.Text;
                MainControllerList.Add(mainController);
                ExecutionOfProgramme clas = new ExecutionOfProgramme();
                //upload the first firmware (calibraion)
                clas.uploadFirmwares(MainControllerList, "controller_test2.ino.with_bootloader.mega.hex", label8);

                var stopwatch = Stopwatch.StartNew();
                Thread.Sleep(6000);
                stopwatch.Stop();
            }

            private void label13_Click(object sender, EventArgs e)
            {

            }
            private void blickPort(string portID)
            {
                using (SerialPort port = new SerialPort(portID))
                {
                    // Set the properties.
                    try
                    {
                        port.BaudRate = 300;
                        port.Parity = Parity.None;
                        port.ReadTimeout = 10;
                        port.StopBits = StopBits.One;

                        // Write a message into the port.
                        port.Open();
                        int i = 0;
                        while (i < 6)
                        {
                            // byte[] bytesToSend = new byte[1] {0x40};

                            port.Close();
                            port.Open();
                            port.Write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
                            port.Close();
                            Thread.Sleep(1000);
                            i++;
                        }

                        port.Close();
                        // Console.WriteLine("Wrote to the port.");
                        //Console.ReadLine();
                    }
                    catch { MessageBox.Show(port.PortName + " port is not conected"); }
                }
            }
            public void lightTheDeviceForRecognize(string buttonName, string command)
            {

                try
                
                    {
                    string portname = buttonName;
                    SerialPort portMain = new SerialPort(portname);
                    if(portname=="0")
                        return;
                    portMain.BaudRate = 115200;
                    portMain.Parity = Parity.None;
                    portMain.ReadTimeout = 100;
                    portMain.StopBits = StopBits.One;
                    try
                    {
                        portMain.Open();
                    }
                    catch { }

                    if (command == "LED_ON")
                    {
                        try
                        {
                            portMain.Write("LED_ON");
                        }
                        catch (Exception ex) { MessageBox.Show("Error"); }
                    }

                    if (command == "LED_OFF")
                    {
                        try
                        {
                            portMain.Write("LED_OFF");
                        }
                        catch (Exception ex) { MessageBox.Show("Error"); }
                    }

                    Thread.Sleep(500);
                    portMain.Close();
                }
                catch (Exception ex) { MessageBox.Show("Error"); }
            }

            private void button40_Click(object sender, EventArgs e)
            {

            }

            public void excecutionErrorSummery()
            {
                String filePath;
                if (Form1.CallibrationDataFolderPath == "None")
                {
                    string workingDir = Directory.GetCurrentDirectory().Replace("WindowsFormsApplication5\\bin\\Release", "").Replace("WindowsFormsApplication5\\bin\\Debug", "");
                    filePath = workingDir + "OffsetValues\\excecutionErrorSummery.txt";
                }
                else filePath = Form1.CallibrationDataFolderPath + "\\excecutionErrorSummery.txt";
                if (!File.Exists(filePath))
                {
                    File.Create(filePath).Close();
                }
                using (StreamWriter sw = File.AppendText(filePath))
                {
                    sw.WriteLine("main controller is disconnected!");
                }
            }

            private void comboBox62_SelectedIndexChanged(object sender, EventArgs e)
            {

            }
    }

    
}
