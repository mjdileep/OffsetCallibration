using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace WindowsFormsApplication5
{
    class ReactiveEnergyExecution
    {
        public int NumberOfAttemptsForData = 4;
        List<string> ReactiveEnergyDatafromRegisters = new List<string>();
        List<string> TestReactiveEnergyCalibrationDatafromRegisters = new List<string>();
        List<SerialPort> portsForReactive = new List<SerialPort>();
        List<SerialPort> TestportsForReactive = new List<SerialPort>();
        static string linecycle;
        static string testLineCycle;
        string ReactiveEnergy;
        string testReactiveEnergy;
        static string energyFromTheMeterForReactive;
        static string testenergyFromTheMeterForReactive;
        static string powerFactor;
        static string powerFactor2;

        public void ReactiveEnergyDoTheProcessAndReadValues(List<Form1.EnableListDeviceInformation> devicesData, string mainport, Label label8, Label label11, Label label12)
        {
            List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;
            //
            //on the meter
           
            try
            {
                SerialPort portMain = new SerialPort(mainport);

                // Set the properties.
                portMain.BaudRate = 115200;
                portMain.Parity = Parity.None;
                portMain.ReadTimeout = 1000;
                portMain.StopBits = StopBits.One;


                // Write a message into the port.
                portMain.Open();
                //change the main
                portMain.Write("CON1,127,255,191,1,\n");
                Thread.Sleep(1000);
                portMain.Write("CON1,127,255,191,0,\n");
                Thread.Sleep(1000);
                portMain.Write("CON1,255,251,191,0,\n");
                Thread.Sleep(1000);
                portMain.Write("CON1,255,251,191,1,\n");
                Thread.Sleep(1000);
              
                portMain.Close();

             
                Thread.Sleep(8000);
           
                // get the correct value from wt1600

                //for display
                try
                {
                    GetdataWT1600 gDisplay = new GetdataWT1600();
                    gDisplay.OpenDevice();
                    string[] WT1600data3 = gDisplay.writeAndtakeAlltheParameters("firstCurrent").Split(',');
                    gDisplay.CloseDevice();

                    label8.Text = "Reactive energry LC1 ";
                    label8.Refresh();

                    label11.Text = WT1600data3[0]; label11.Refresh();
                    label12.Text = WT1600data3[1]; label12.Refresh();
                    powerFactor  = WT1600data3[2];
                }
                catch {
                    GetdataWT1600 gDisplay = new GetdataWT1600();
                    gDisplay.CloseDevice(); }
                //////for display
                var stopwatchbeforevolt34 = Stopwatch.StartNew();
                Thread.Sleep(3000);
                stopwatchbeforevolt34.Stop();

                GetdataWT1600Energy getActiveEnergy = new GetdataWT1600Energy();
                getActiveEnergy.OpenDevice();
                string time = "0,1,40";
                getActiveEnergy.startIntegration(time,"high");


                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        label8.Text = "Reactive energry LC1 " + (i + 1) + "/" + deviceDatalist.Count + "(Total devices)";
                        label8.Refresh();
                        //do whatever need to do to these ports one by one

                        SerialPort port = new SerialPort();
                        port.PortName = deviceDatalist[i].Port;
                        int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[i].Ct);

                        // Set the properties.
                        port.BaudRate = 115200;
                        port.Parity = Parity.None;
                        // port.ReadTimeout = 10000;
                        port.StopBits = StopBits.One;
                        //port.RtsEnable = true;
                        portsForReactive.Add(port);
                        // Write a message into the port.
                      //  Thread.Sleep(32000);
                        Console.WriteLine("waited");
                        try
                        {
                            port.Open();
                            Console.WriteLine("success");
                        }
                        catch (Exception ex) { 
                            Console.WriteLine(ex); }
                      
                       
                        // port.DiscardOutBuffer();

                        var stopwatchPhCal = Stopwatch.StartNew();
                        Thread.Sleep(1000);
                        stopwatchPhCal.Stop();
                        //write the messages ,can use relevent CT and deviceId




                        int lineCircles = 10000;
                        linecycle = lineCircles.ToString();
                        port.ReadExisting();//just to make sure indata buffer is cleaned
                        string command = "RP_CAL," + lineCircles + ",\n";
                        port.Write(command);
                        Console.WriteLine(command + "is written");
                    }
                    catch { continue; }
                }
                //wait till lineCycle complets and collect data
                double secondsDouble = Convert.ToDouble("10000");

                int linecircles = (int)(Math.Ceiling(secondsDouble));
                ///wait
                ///

               
                Thread.Sleep(linecircles * 10 + 6000);
             
                //////////////////////////////////////////////////////////////////////////////////////////////
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    //  string portnametemp = "port" + i;
                    // SerialPort port = new SerialPort(deviceDatalist[i].Port);
                    try
                    {
                        int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[i].Ct);
                        int temp1 = 0;
                        // portsForReactive[i].DiscardInBuffer();
                        while (ReactiveEnergy == null || ReactiveEnergy == "")
                        {

                            if (temp1 >= NumberOfAttemptsForData) { break; }
                            var stopwatch = Stopwatch.StartNew();
                            Thread.Sleep(2000);
                            stopwatch.Stop();
                            ReactiveEnergy = portsForReactive[i].ReadExisting();
                            Console.WriteLine(temp1 + "=" + ReactiveEnergy);
                            temp1++;
                            
                        }
                        if (ReactiveEnergy == "" || ReactiveEnergy == null)
                            this.excecutionErrorSummery(Int32.Parse(deviceDatalist[i].DeviceId), "Initial reactive energy");
                        ReactiveEnergyDatafromRegisters.Add(ReactiveEnergy);
                        ReactiveEnergy = null;
                        //when device gives the data ,stop the meter too


                        portsForReactive[i].DiscardInBuffer();
                        portsForReactive[i].DiscardOutBuffer();
                        portsForReactive[i].Close();
                    }
                    catch { continue; }
                }

                //finally finish the integration on meter and take the value
                string[] WT1600data = getActiveEnergy.returnEnergy("Energynormal").Split(',');
              //  double tempmeter = double.Parse(WT1600data[0]);
               // double tempPower = (Math.Acos((double)(decimal.Parse(powerFactor))));
                //double tempFinal=Math.Tan(Math.Acos((double)(decimal.Parse(powerFactor))));
                energyFromTheMeterForReactive = ((double.Parse(WT1600data[0])) * Math.Tan(Math.Acos((double)(decimal.Parse(powerFactor))))).ToString();
                //we obtained active energy convert it to reactive
                


                getActiveEnergy.CloseDevice();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

        }
        public void excecutionErrorSummery(int id, string field)
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
                sw.WriteLine("Field :" + field + " from device :" + id + " was not recieved!");
            }
        }

        internal void DotheReactiveEnergyCalculations(List<Form1.EnableListDeviceInformation> devicesData, string mainControllerPort, Label label11,Label label12, Label label8 )
        {
            try
            {
                var ReactiveEnergyObjects = new List<ReactiveEnergylistfromgatherdData>();
                List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        ReactiveEnergylistfromgatherdData temp = new ReactiveEnergylistfromgatherdData();
                        temp.DeviceId = deviceDatalist[i].DeviceId;
                        string[] phases = ReactiveEnergyDatafromRegisters[i].Split(',');
                        temp.phaseAenergy = phases[1];
                        temp.phaseBenergy = phases[2];
                        temp.phaseCenergy = phases[3];
                        temp.Wt1600MeasuredReactiveEnergy = energyFromTheMeterForReactive;
                        temp.lineCircle = linecycle;
                        temp.CT = deviceDatalist[i].Ct;
                        temp.port = deviceDatalist[i].Port;
                        ReactiveEnergyObjects.Add(temp);
                    }
                    catch { continue; }


                }

                SerialPort port2 = new SerialPort(mainControllerPort);

                // Set the properties.
                port2.BaudRate = 115200;
                port2.Parity = Parity.None;
                port2.ReadTimeout = 6000;
                port2.StopBits = StopBits.One;

                // Write a message into the port.
               
                port2.Open();
                port2.Write("CON1,255,251,191,1,\n"); Thread.Sleep(1000);
                port2.Write("CON1,255,251,191,0,\n"); Thread.Sleep(1000);
                port2.Write("CON1,255,127,191,0,\n"); Thread.Sleep(1000);
                port2.Write("CON1,255,127,191,1,\n"); Thread.Sleep(1000);
              //  port2.Write("CON1,255,127,191,1,\n"); //test energy change
                port2.Close();
              
                Thread.Sleep(8000);
            

                //change the main here,start the meter
                label8.Text = "Reactive energry LC2 ";
                label8.Refresh();
                //for display
                try
                {
                    GetdataWT1600 gDisplay = new GetdataWT1600();
                    gDisplay.OpenDevice();
                    string[] WT1600data3 = gDisplay.writeAndtakeAlltheParameters("energynormal").Split(',');
                    gDisplay.CloseDevice();
                    label11.Text = WT1600data3[0]; label11.Refresh();
                    label12.Text = WT1600data3[1]; label12.Refresh();
                    powerFactor2 = WT1600data3[2];
                }
                catch { }
            
                /////for display

                GetdataWT1600Energy getActiveEnergy = new GetdataWT1600Energy();
                getActiveEnergy.OpenDevice();
                string time2 = "0,0,2";
                getActiveEnergy.startIntegration(time2,"low");


                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        label8.Text = "Reactive energry LC2 " + (i + 1) + "/" + deviceDatalist.Count + "(Total devices)";
                        label8.Refresh();
                        //do whatever need to do to these ports one by one
                        SerialPort port = new SerialPort(deviceDatalist[i].Port);
                        int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[i].Ct);

                        // Set the properties.
                        port.BaudRate = 115200;
                        port.Parity = Parity.None;
                       // port.ReadTimeout = 2000;
                        port.StopBits = StopBits.One;
                        //port.RtsEnable = true;

                        // Write a message into the port.


                        TestportsForReactive.Add(port);

                      //  Thread.Sleep(30000);
                        try
                        {
                            port.Open();
                        }
                        catch{Console.WriteLine("error");}
                        Thread.Sleep(1000);

                        //write the messages ,can use relevent CT and deviceId
                        int lineCircles =200;
                        testLineCycle = lineCircles.ToString();
                        port.ReadExisting();
                        Thread.Sleep(1000);
                        string command ="RP_CAL," + lineCircles + ",";
                        port.Write(command);
                        Console.WriteLine(command + "written");
                    }
                    catch { continue; }
                }

                double secondsDouble = Convert.ToDouble("200");

                int linecircles = (int)(Math.Ceiling(secondsDouble));
                Console.WriteLine(linecircles);
               
                Thread.Sleep(linecircles * 10 + 6000);
                Console.WriteLine(linecircles * 1000 + 4000 +" waited");

                //////////////////////////////////////////////////////////////////////////////////////////////
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        //SerialPort port = new SerialPort(deviceDatalist[i].Port);
                        int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[i].Ct);

                        int temp1 = 0;
                        while (testReactiveEnergy == null || testReactiveEnergy == "")
                        {

                            if (temp1 >= NumberOfAttemptsForData) { break; }

                            Thread.Sleep(2000);

                            testReactiveEnergy = TestportsForReactive[i].ReadExisting();
                            string tempstr = testReactiveEnergy;


                            if (tempstr != "")
                            {
                                if (tempstr.ToLowerInvariant().IndexOf('\n') == -1)
                                {
                                    Console.WriteLine("full length didn't come");
                                    var stop = Stopwatch.StartNew();
                                    Thread.Sleep(3000);
                                    stop.Stop();
                                    testReactiveEnergy = TestportsForReactive[i].ReadExisting();
                                    testReactiveEnergy = tempstr + testReactiveEnergy;
                                }
                            }

                            
                            Console.WriteLine(temp1 + "=" + testReactiveEnergy);
                            temp1++;
                        }
                        if (testReactiveEnergy == "" || testReactiveEnergy == null)
                            this.excecutionErrorSummery(Int32.Parse(deviceDatalist[i].DeviceId), "Test reactive energy");
                        TestReactiveEnergyCalibrationDatafromRegisters.Add(testReactiveEnergy);
                        testReactiveEnergy = null;
                      
                            TestportsForReactive[i].DiscardInBuffer();
                            TestportsForReactive[i].DiscardOutBuffer();
                            TestportsForReactive[i].Close();
                    }
                    catch { continue; }
                }

                string[] WT1600data = getActiveEnergy.returnEnergy("testActive").Split(',');
                testenergyFromTheMeterForReactive = (((double.Parse(WT1600data[0]))) * Math.Tan(Math.Acos((double)(decimal.Parse(powerFactor2))))).ToString();
                ;

                getActiveEnergy.CloseDevice();
                ///////////////////////////////////////////////////////////////////////////////////////


                var stopwatchcurrent = Stopwatch.StartNew();
                Thread.Sleep(2000);
                stopwatchcurrent.Stop();
                TestReactiveEnergyCalculations(devicesData, ReactiveEnergyObjects);
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void TestReactiveEnergyCalculations(List<Form1.EnableListDeviceInformation> devicesData, List<ReactiveEnergylistfromgatherdData> ReactiveEnergyDataObjects)
        {
            var TestEnergyDataObjects = new List<TestReactiveEnergylistfromgatherdData>();
            List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;
            for (int k = 0; k < deviceDatalist.Count; k++)
            {
                try
                {
                    TestReactiveEnergylistfromgatherdData temp = new TestReactiveEnergylistfromgatherdData();
                    temp.DeviceId = deviceDatalist[k].DeviceId;
                    string[] phases2 = TestReactiveEnergyCalibrationDatafromRegisters[k].Split(',');
                    temp.TestphaseAenergy = phases2[1];
                    temp.TestphaseBenergy = phases2[2];
                    temp.TestphaseCenergy = phases2[3];

                    //get the correct value from wt1600
                    temp.TestWt1600MeasuredEnergy = testenergyFromTheMeterForReactive;
                    temp.TestlineCircle = testLineCycle;




                    TestEnergyDataObjects.Add(temp);
                }
                catch { continue; }


            }
            Equations eqClass = new Equations();
            eqClass.ReactiveEneryCalibarionCalculations(TestEnergyDataObjects, ReactiveEnergyDataObjects);
        }

        public class ReactiveEnergylistfromgatherdData
        {
            public string CT{ get; set; }
            public string DeviceId { get; set; }
            public string Wt1600MeasuredReactiveEnergy { get; set; }
            public string phaseAenergy { get; set; }
            public string phaseBenergy { get; set; }
            public string phaseCenergy { get; set; }
            public string lineCircle { get; set; }
            public string port { get; set; }

        }

        public class TestReactiveEnergylistfromgatherdData
        {
            public string DeviceId { get; set; }
            public string TestWt1600MeasuredEnergy { get; set; }
            public string TestphaseAenergy { get; set; }
            public string TestphaseBenergy { get; set; }
            public string TestphaseCenergy { get; set; }
            public string TestlineCircle { get; set; }
        }



    }

   
}
