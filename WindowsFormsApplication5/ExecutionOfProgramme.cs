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
using System.IO;



namespace WindowsFormsApplication5
{
    class ExecutionOfProgramme
    {
        public int NumberOfAttemptsForData = 13;
        List<string> CurrentCalibrationDatafromRegisters = new List<string>();
        List<string> TestCurrentCalibrationDatafromRegisters = new List<string>();
        
        List<string> VoltageCalibrationDatafromRegisters = new List<string>();
        List<string> TestVoltageCalibrationDatafromRegisters = new List<string>();

        List<string> PhaseCalibrationDatafromRegisters = new List<string>();
        List<string> TestPhaseCalibrationDatafromRegisters = new List<string>();

        List<string> ActiveEnergyDatafromRegisters = new List<string>();
        List<string> TestActiveEnergyCalibrationDatafromRegisters = new List<string>();

        string indatacurrent; string indataTestcurrent; string indataVoltage; string indataTestVoltage; string indataPhase; string indataTestphase;
        
        string activeEnergy; string testActiveEnergy;
        
        static string energyFromTheMeterForActive;
        static string testenergyFromTheMeterForActive;

        static string linecycle;
        static string testLineCycle;

        static string currentPortData;

       


        List<SerialPort> ports = new List<SerialPort>();
        List<SerialPort> Testports = new List<SerialPort>();

        public void maincontroller(string port)
        { //open the main controllers port and wtrite to it
            try
            {
                SerialPort portMain = new SerialPort(port);
                //  int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                // int CT = Int32.Parse(deviceDatalist[i].Ct);

                // Set the properties.
                portMain.BaudRate = 115200;
                portMain.Parity = Parity.None;
                portMain.ReadTimeout = 10;
                portMain.StopBits = StopBits.One;

            }
            catch { Console.Write("error"); }
        }
        
        
        
        //firmware uploader
        public void uploadFirmwares(List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> devicesData,string firmwareName,Label label8) {

            List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;
            try
            {
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        label8.Text = "Uploading Firmware" + (i + 1) + "/" + deviceDatalist.Count;
                        label8.Refresh();
                        //do whatever need to do to these ports one by one
                        SerialPort port = new SerialPort(deviceDatalist[i].Port);
                        //  int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        //  int CT = Int32.Parse(deviceDatalist[i].Ct);

                        // Set the properties.
                        port.BaudRate = 115200;
                        port.Parity = Parity.None;
                        port.ReadTimeout = 10;
                        port.StopBits = StopBits.One;

                        //upload the firmares oneby one
                        port.PortName = deviceDatalist[i].Port;

                        string workingDirectory = Directory.GetCurrentDirectory();
                        workingDirectory=workingDirectory.Replace("\\", "/");
                        workingDirectory = workingDirectory.Replace("/WindowsFormsApplication5/bin/Release", "/");
                        workingDirectory = workingDirectory.Replace("/WindowsFormsApplication5/bin/Debug", "/");
                        string Filpath = workingDirectory + "AVRDUDEpro";
                        
                        string Command = Filpath + "/epro" + " -C " + Filpath + "/epro.conf" + " -v -patmega2560 -cwiring -" + "P" + port.PortName + " -b115200 -D -Uflash:w:" + Filpath + "/" + firmwareName + ":i";

                        ProcessStartInfo ProcessInfo;
                        Process Process;

                        ProcessInfo = new ProcessStartInfo("cmd.exe", "/K " + Command);
                        ProcessInfo.CreateNoWindow = true;
                        ProcessInfo.UseShellExecute = false;
                        ProcessInfo.CreateNoWindow = true;
                        Process = Process.Start(ProcessInfo); //something tricky*************************************************************
                        int millicount = 0;
                        while (!Process.HasExited&&millicount<10)
                        {
                            Thread.Sleep(1);
                            millicount++;
                             
                        }
                            port.Close();

                    }
                    catch { continue; }
                }
                label8.Text = "Firmware is uploaded." ;
                label8.Refresh();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

           
        }

        //wtrite commands to registers and read values
        public void DotheProcessAndReadvalues(List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> devicesData,Label label8,Label label11,Label label12)
        { 
            //in this list got the port number .Have to open that port number and send the string messages

            List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;

            try
            {
                string status = "Initial current" ;
                label8.Text = status;
                label8.Refresh();

                try
                {
                    GetdataWT1600 getdata = new GetdataWT1600();
                    getdata.OpenDevice();
                    string[] WT1600data = getdata.writeAndtakeAlltheParameters("firstCurrent").Split(',');
                    getdata.CloseDevice();

                    label11.Text = WT1600data[0];
                    label12.Text = WT1600data[1];
                    label11.Refresh(); label12.Refresh();
                }
                catch { }
                List<SerialPort> openList = new List<SerialPort>();
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    SerialPort port = new SerialPort(deviceDatalist[i].Port);
                    int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                    int CT = Int32.Parse(deviceDatalist[i].Ct);

                    // Set the properties.
                    port.BaudRate = 115200;
                    port.Parity = Parity.None;
                    // port.ReadTimeout = 1000;
                    port.StopBits = StopBits.One;
                    // //port.RtsEnable = true;
                    // port.Close();
                    // Write a message into the port.
                    openList.Add(port);
                    try
                    {
                        port.Open();

                        var stopwatchcCal = Stopwatch.StartNew();
                        Thread.Sleep(1000);
                        stopwatchcCal.Stop();
                        //write the messages ,can use relevent CT and deviceId
                        //write the messages ,can use relevent CT and deviceId
                        port.ReadExisting();
                        port.Write("C_CAL");
                    }
                    catch (UnauthorizedAccessException uAE)
                    {
                        MessageBox.Show("Please re-plug the device cable/cables and try!");
                        System.Environment.Exit(1);
                    }
                    catch (Exception uAE)
                    {
                        this.excecutionErrorSummery(deviceDatalist[i].Port);
                        deviceDatalist.RemoveAt(i);
                        continue;
                    }
                }
                
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        string status2 = "Initial current : " + (i + 1) + "/" + deviceDatalist.Count+"(Total devices)";
                        label8.Text = status2;
                        label8.Refresh();
                        //do whatever need to do to these ports one by one
                        SerialPort port = openList[i];
                        int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[i].Ct);

                       // //port.RtsEnable = true;
                        // port.Close();
                        // Write a message into the port.
                        Thread.Sleep(1000);

                        // port.Write(Command);
                        //multi tread this after here

                        string indatacurrent = null; int temp1 = 0;
                        while (indatacurrent == null || indatacurrent == "")
                        {
                            // Console.WriteLine("time =" + temp1);
                            if (temp1 == NumberOfAttemptsForData) { break; }
                            if (temp1 > NumberOfAttemptsForData/2)
                            {
                                port.DiscardOutBuffer();
                                port.Write("C_CAL");

                                Thread.Sleep(3000);


                            }

                            Thread.Sleep(1000);

                            indatacurrent = port.ReadExisting();
                            string tempstr = indatacurrent;


                            if (tempstr != null)
                            {
                                if (tempstr.ToLowerInvariant().IndexOf('\n') == -1)
                                {
                                    Thread.Sleep(3000);

                                    indatacurrent = port.ReadExisting();
                                    indatacurrent = tempstr + indataTestphase;
                                }
                            }

                            Console.WriteLine(temp1 + "=" + indataTestphase);
                            temp1++;

                        }
                        if (indatacurrent == "" || indatacurrent == null)
                            this.excecutionErrorSummery(Int32.Parse(deviceDatalist[i].DeviceId), "Initial current");
                        CurrentCalibrationDatafromRegisters.Add(indatacurrent);
                        port.DiscardInBuffer();
                        port.DiscardOutBuffer();
                        port.Close();
                    }
                    catch (Exception ex) { Console.WriteLine(   ex); continue; }

                    
                }
                //change the main here before call to the test current
                // string mainControllerPort = comboBox61.Text;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

          

                
        }
        
        public void doTheCurrentCalculations(List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> devicesData, string mainPort, Label label11, Label label12, Label label8)
        {
            try
            {
                var CurrentDataObjects = new List<currentlistfromgatherdData>();
                List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;

                GetdataWT1600 getdata2 = new GetdataWT1600();
                getdata2.OpenDevice();
                string[] WT1600data2 = getdata2.writeAndtakeAlltheParameters("firstCurrent").Split(',');
                getdata2.CloseDevice();
                
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                   
                    try
                    {
                        currentlistfromgatherdData temp = new currentlistfromgatherdData();
                        temp.DeviceId = deviceDatalist[i].DeviceId;
                        string[] phases = CurrentCalibrationDatafromRegisters[i].Split(',');
                        temp.PhaseA = phases[1];
                        temp.PhaseB = phases[2];
                        temp.PhaseC = phases[3];
                        temp.CT = deviceDatalist[i].Ct;
                        temp.port = deviceDatalist[i].Port;
                        //get the correct value from wt1600
                        temp.Wt1600Readcurrent = WT1600data2[0];

                        // getdata.CloseDevice();

                        CurrentDataObjects.Add(temp);

                    }
                    catch (Exception ex) { Console.WriteLine(ex); continue; }

                }

                SerialPort port2 = new SerialPort(mainPort);

                // Set the properties.
                port2.BaudRate = 115200;
                port2.Parity = Parity.None;
              //  port2.ReadTimeout = 1000;
                port2.StopBits = StopBits.One;

                // Write a message into the port.
                try
                {
                    port2.Open();
                    port2.Write("CON1,127,255,191,1,\n");   //just off the main
                    Thread.Sleep(1000);
                    port2.Write("CON1,127,255,191,0,\n");   //off the trioc 
                    Thread.Sleep(1000);
                    port2.Write("CON1,159,255,191,0,\n");   //change the loads
                    Thread.Sleep(1000);
                    port2.Write("CON1,159,255,191,1,\n");    //on the trioc
                    Thread.Sleep(1000);
                    port2.Write("CON1,159,255,191,1,\n");   //on the main
                    port2.Close();
                }
                catch (Exception ex)
                {
                    this.excecutionErrorSummery("Can't connect to main controller!");
                    System.Environment.Exit(1);
                }
                var stop3 = Stopwatch.StartNew();
                Thread.Sleep(5000);
                stop3.Stop();

                string status2 = "Test current ";
                label8.Text = status2;
                label8.Refresh();

                try
                {
                    GetdataWT1600 getdata = new GetdataWT1600();
                    getdata.OpenDevice();
                    string[] WT1600data = getdata.writeAndtakeAlltheParameters("TestCurrent").Split(',');
                    getdata.CloseDevice();

                    label11.Text = WT1600data[0];
                    label12.Text = WT1600data[1];
                    label11.Refresh(); label12.Refresh();
                }
                catch { }
                List<SerialPort> openList = new List<SerialPort>();
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    SerialPort port = new SerialPort(deviceDatalist[i].Port);
                    int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                    int CT = Int32.Parse(deviceDatalist[i].Ct);

                    // Set the properties.
                    port.BaudRate = 115200;
                    port.Parity = Parity.None;
                    // port.ReadTimeout = 1000;
                    port.StopBits = StopBits.One;
                    // //port.RtsEnable = true;
                    // port.Close();
                    // Write a message into the port.
                    openList.Add(port);
                    try
                    {
                        port.Open();
                        var stopwatchcCal = Stopwatch.StartNew();
                        Thread.Sleep(1000);
                        stopwatchcCal.Stop();
                        //write the messages ,can use relevent CT and deviceId
                        //write the messages ,can use relevent CT and deviceId
                        port.ReadExisting();
                        port.Write("C_CAL");
                    }
                    catch (Exception uAE)
                    {
                        this.excecutionErrorSummery(deviceDatalist[i].Port);
                        deviceDatalist.RemoveAt(i);
                        continue;
                    }
                    
                }

                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    string status = "Test current : " + (i + 1) + "/" + deviceDatalist.Count + "(Total devices)";
                    label8.Text = status;
                    label8.Refresh();
                    try
                    {
                        //do whatever need to do to these ports one by one
                        SerialPort port = openList[i];
                        int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[i].Ct);

                        Thread.Sleep(1000);

                        string indataTestcurrent = null; int temp1 = 0;
                        while (indataTestcurrent == null || indataTestcurrent == "")
                        {
                            // Console.WriteLine("time =" + temp1);
                            if (temp1 == NumberOfAttemptsForData) { break; }
                            if (temp1 > NumberOfAttemptsForData/2)
                            {
                                port.DiscardOutBuffer();
                                port.Write("C_CAL");
                                var s = Stopwatch.StartNew();
                                Thread.Sleep(3000);
                                s.Stop();

                            }

                            Thread.Sleep(1000);

                            indataTestcurrent = port.ReadExisting();
                            string tempstr = indataTestcurrent;


                            if (tempstr != null)
                            {
                                if (tempstr.ToLowerInvariant().IndexOf('\n') == -1)
                                {
                                    //  Console.WriteLine("full length didn't come");

                                    Thread.Sleep(3000);

                                    indataTestcurrent = port.ReadExisting();
                                    indataTestcurrent = tempstr + indataTestcurrent;
                                }
                            }

                            // Console.WriteLine(j + "=" + indataTestphase);
                            temp1++;

                        }
                        if (indataTestcurrent == "" || indataTestcurrent == null)
                            this.excecutionErrorSummery(Int32.Parse(deviceDatalist[i].DeviceId), "Test current");
                        TestCurrentCalibrationDatafromRegisters.Add(indataTestcurrent);
                        port.DiscardInBuffer();
                        port.DiscardOutBuffer();
                        port.Close();
                    }
                    catch (Exception ex) { Console.WriteLine(ex); continue; }
                    //put here the statue update

                }

                var stopwatchcurrent = Stopwatch.StartNew();
                Thread.Sleep(3000);
                stopwatchcurrent.Stop();
                TestCurrentCalculations(devicesData, CurrentDataObjects, label11, label12 ,label8);
              

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void  TestCurrentCalculations(List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> devicesData, List<currentlistfromgatherdData> currentlist, Label label11, Label label12,Label label8)
        {
            try
            {
                var TestCurrentDataObjects = new List<testcurrentfromgatherdata>();
                //get the correct value from wt1600
                GetdataWT1600 getdata = new GetdataWT1600();
                getdata.OpenDevice();
                string[] WT1600data = getdata.writeAndtakeAlltheParameters("testcurrentnormal").Split(',');
                getdata.CloseDevice();

                List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {

                        testcurrentfromgatherdata temp = new testcurrentfromgatherdata();
                        temp.DeviceId = deviceDatalist[i].DeviceId;
                        string[] phases = TestCurrentCalibrationDatafromRegisters[i].Split(',');
                        temp.testPhaseA = phases[1];
                        temp.testPhaseB = phases[2];
                        temp.testPhaseC = phases[3];

                        
                        temp.testWt1600Readcurrent = WT1600data[0];
                        // getdata.CloseDevice();


                        TestCurrentDataObjects.Add(temp);
                    }
                    catch (Exception ex) { Console.WriteLine(ex); continue; }


                }
                Equations eqClass = new Equations();
                eqClass.currentCalibarionCalculations(TestCurrentDataObjects, currentlist);
                label11.Text = " "; label11.Refresh();
                label12.Text = " "; label12.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }




      //voltage
        public void VoltageDotheProcessAndReadvalues(List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> devicesData,Label label8,Label label11,Label label12)
        {
            //in this list got the port number .Have to open that port number and send the string messages
           
                List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;

                try
                {

                    try
                    {
                        GetdataWT1600 getdata = new GetdataWT1600();
                        getdata.OpenDevice();
                        string[] WT1600data = getdata.writeAndtakeAlltheParameters("firstCurrent").Split(',');
                        getdata.CloseDevice();

                        label11.Text = WT1600data[0]; label11.Refresh();
                        label12.Text = WT1600data[1]; label12.Refresh();
                    }
                    catch { }
                    List<SerialPort> openList = new List<SerialPort>();
                    for (int i = 0; i < deviceDatalist.Count; i++)
                    {
                        SerialPort port = new SerialPort(deviceDatalist[i].Port);
                        int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[i].Ct);

                        // Set the properties.
                        port.BaudRate = 115200;
                        port.Parity = Parity.None;
                        // port.ReadTimeout = 1000;
                        port.StopBits = StopBits.One;
                        // //port.RtsEnable = true;
                        // port.Close();
                        openList.Add(port);
                        try
                        {
                            port.Open();
                            var stopwatchcCal = Stopwatch.StartNew();
                            Thread.Sleep(1000);
                            stopwatchcCal.Stop();
                            //write the messages ,can use relevent CT and deviceId
                            //write the messages ,can use relevent CT and deviceId
                            port.ReadExisting();
                            port.Write("V_CAL");
                        }
                        catch (Exception ex)
                        {
                            this.excecutionErrorSummery(deviceDatalist[i].Port);
                            deviceDatalist.RemoveAt(i);
                            continue;
                        }
                    }
                    for (int i = 0; i < deviceDatalist.Count; i++)
                    {
                        try
                        {
                            string status = "Initial voltage : " + (i + 1) + "/" + deviceDatalist.Count + "(Total devices)";
                            label8.Text = status;
                            label8.Refresh();



                            //do whatever need to do to these ports one by one
                            SerialPort port = openList[i];
                            int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                            int CT = Int32.Parse(deviceDatalist[i].Ct);

                            // Set the properties.
                           
                            Thread.Sleep(1000);

                            //multi thread this after here
                            string indataVoltage = null; int temp1 = 0;
                            while (indataVoltage == null || indataVoltage == "")
                            {
                                Console.WriteLine("time =" + temp1);
                                if (temp1 == this.NumberOfAttemptsForData) {  break; }
                                if (temp1 > this.NumberOfAttemptsForData/2)
                                {
                                    port.DiscardOutBuffer();
                                    port.Write("V_CAL");
                                    var s = Stopwatch.StartNew();
                                    Thread.Sleep(3000);
                                    s.Stop();

                                }

                                Thread.Sleep(1000);

                                indataVoltage = port.ReadExisting();
                                string tempstr = indataVoltage;


                                if (tempstr != null)
                                {
                                    if (tempstr.ToLowerInvariant().IndexOf('\n') == -1)
                                    {
                                        Console.WriteLine("full length didn't come");
                                        var stop = Stopwatch.StartNew();
                                        Thread.Sleep(3000);
                                        stop.Stop();
                                        indataVoltage = port.ReadExisting();
                                        indataVoltage = tempstr + indataVoltage;
                                    }
                                }

                                //Console.WriteLine(j + "=" + indataVoltage);
                                temp1++;

                            }

                            if (indataVoltage == "" || indataVoltage == null)
                                this.excecutionErrorSummery(Int32.Parse(deviceDatalist[i].DeviceId), "Intial voltage");
                            VoltageCalibrationDatafromRegisters.Add(indataVoltage);
                            port.DiscardInBuffer();
                            port.DiscardOutBuffer();
                            port.Close();
                        }
                        catch (Exception ex) { Console.WriteLine(ex); continue; }
                    }

                }
              
            
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }
        public void VoltageDotheProcessAndReadvalues(List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> devicesData, Label label8, Label label11, Label label12,bool overloard)
        {
            //in this list got the port number .Have to open that port number and send the string messages

            List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;

                try
                {
                    GetdataWT1600 getdata = new GetdataWT1600();
                    getdata.OpenDevice();
                    string[] WT1600data = getdata.writeAndtakeAlltheParameters("firstCurrent").Split(',');
                    getdata.CloseDevice();
                    label11.Text = WT1600data[0]; label11.Refresh();
                    label12.Text = WT1600data[1]; label12.Refresh();
                }
                catch { }
                ///request for data
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                        string status = "Reqesting initial voltage : " + (i + 1) + "/" + deviceDatalist.Count + "(Total devices)";
                        label8.Text = status;
                        label8.Refresh();
                        this.askForData(deviceDatalist[i].Port, "V_CAL");
                }
                //read from devices
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    string status = "Reading initial voltage : " + (i + 1) + "/" + deviceDatalist.Count + "(Total devices)";
                    label8.Text = status;
                    label8.Refresh();
                    string buffer=this.getDataFromPorts(deviceDatalist[i].Port);
                    if (buffer == null) MessageBox.Show("No data from port: " + deviceDatalist[i].Port);
                    else
                        VoltageCalibrationDatafromRegisters.Add(buffer);
                }



        }
        public string getDataFromPorts(string portName)///will rerurn null if catch an exception
        {
            SerialPort port = new SerialPort(portName);
            string buffer=null;
            try
            {
                port.BaudRate = 115200;
                port.Parity = Parity.Even;
                port.StopBits = StopBits.One;
                port.Open();
                buffer=port.ReadExisting();
                if (buffer != null)
                {
                    int count = 0;
                    while(buffer.ToLowerInvariant().IndexOf('\n') == -1 &&count<3)
                    {
                        count++;
                        Thread.Sleep(1000);
                        Console.WriteLine("Full length didn't come, getting the rest....");
                        indataVoltage = buffer + port.ReadExisting();
                    }
                    if (count > 3)
                    {
                        Console.WriteLine("Port took long time to respond, exiting....");
                        buffer = null;
                    }
                }
                port.DiscardInBuffer();
                port.DiscardOutBuffer();
                port.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while reading from port: " + portName + "\nError: " + ex.ToString());
            }
            return buffer;
        }
        public void askForData(string portName, string cmd)   
        {
                SerialPort port=new SerialPort(portName);
                try
                {
                    port.BaudRate = 115200;
                    port.Parity = Parity.Even;
                    port.StopBits = StopBits.One;
                    port.Open();
                    port.ReadExisting();
                    port.Write(cmd);
                    port.DiscardInBuffer();
                    port.DiscardOutBuffer();
                    port.Close();
                   
                }
                catch (Exception ex) {
                    MessageBox.Show("Error while writing to port: " + portName + " on '" + cmd + "' command!" + "\nError: " + ex.ToString());
                }
            
        }

        public void doTheVoltageCalculations(List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> devicesData, string mainport, Label label11, Label label12, Label label8)
        {
            try
            {
                GetdataWT1600 getdata = new GetdataWT1600();
                getdata.OpenDevice();
                string[] WT1600data = getdata.writeAndtakeAlltheParametersForVoltage().Split(',');
                getdata.CloseDevice();

                var VoltageDataObjects = new List<voltagelistfromgatherdData>();
                List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        voltagelistfromgatherdData temp = new voltagelistfromgatherdData();
                        temp.DeviceId = deviceDatalist[i].DeviceId;
                        string[] phases = VoltageCalibrationDatafromRegisters[i].Split(',');
                        temp.PhaseA = phases[1];
                        temp.PhaseB = phases[2];
                        temp.PhaseC = phases[3];
                        temp.port = deviceDatalist[i].Port;
                        //get the correct value from wt1600
                        temp.Wt1600ReadVoltage = WT1600data[1];

                        VoltageDataObjects.Add(temp);
                    }
                    catch (Exception ex) { Console.WriteLine(ex); continue; }
                   
                }

                SerialPort port2 = new SerialPort(mainport);

                // Set the properties.
                port2.BaudRate = 115200;
                port2.Parity = Parity.None;
               // port2.ReadTimeout = 1000;
                port2.StopBits = StopBits.One;
                port2.RtsEnable = true;
                
             
                // Write a message into the port.
                try
                {
                    port2.Open();
                    port2.Write("CON1,127,255,254,1,\n");
                    Thread.Sleep(1000);
                    port2.Write("CON1,127,255,254,0,\n");
                    Thread.Sleep(1000);
                    port2.Write("CON1,127,255,191,0,\n");  //test voltage change
                    Thread.Sleep(1000);
                    port2.Write("CON1,127,255,191,1,\n");
                    Thread.Sleep(1000);
                    port2.Write("CON1,127,255,191,1,\n");   //test voltage change with main
                    port2.Close();
                }
                catch (Exception ex)
                {
                    this.excecutionErrorSummery("Can't connect to main controller!");
                    System.Environment.Exit(1);
                }
                Thread.Sleep(4000);

                try
                {
                    GetdataWT1600 getdata3 = new GetdataWT1600();
                    getdata3.OpenDevice();
                    string[] WT1600data2 = getdata3.writeAndtakeAlltheParameters("firstCurrent").Split(',');
                    getdata3.CloseDevice();

                    label11.Text = WT1600data2[0];
                    label12.Text = WT1600data2[1];
                    label11.Refresh(); label12.Refresh();
                }
                catch { }
                //change the main here,then call testcurrent
                List<SerialPort> openList = new List<SerialPort>();
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    SerialPort port = new SerialPort(deviceDatalist[i].Port);
                    int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                    int CT = Int32.Parse(deviceDatalist[i].Ct);

                    // Set the properties.
                    port.BaudRate = 115200;
                    port.Parity = Parity.None;
                    // port.ReadTimeout = 1000;
                    port.StopBits = StopBits.One;
                    // //port.RtsEnable = true;
                    // port.Close();
                    openList.Add(port);
                    try
                    {
                        port.Open();
                        var stopwatchcCal = Stopwatch.StartNew();
                        Thread.Sleep(1000);
                        stopwatchcCal.Stop();
                        //write the messages ,can use relevent CT and deviceId
                        //write the messages ,can use relevent CT and deviceId
                        port.ReadExisting();
                        port.Write("V_CAL");
                    }
                    catch (Exception ex)
                    {
                        this.excecutionErrorSummery(deviceDatalist[i].Port);
                        deviceDatalist.RemoveAt(i);
                        continue;
                    }
                }
                for (int j = 0; j < deviceDatalist.Count; j++)
                {
                    try
                    {
                        string status = "Test Voltage : " + (j + 1) + "/" + deviceDatalist.Count + "(Total devices)";
                        label8.Text = status;
                        label8.Refresh();
                       
                        //do whatever need to do to these ports one by one
                        SerialPort port = openList[j];
                        int deviceId = Int32.Parse(deviceDatalist[j].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[j].Ct);

                        Thread.Sleep(1000);

                        // port.Write(Command);
                        string indataTestVoltage = null; int temp1 = 0;
                        while (indataTestVoltage == null || indataTestVoltage == "")
                        {
                            Console.WriteLine("time =" + temp1);
                            if (temp1 == NumberOfAttemptsForData) {  break; }
                            if (temp1 > NumberOfAttemptsForData/2)
                            {
                                port.DiscardOutBuffer();
                                port.Write("V_CAL");
                                var s = Stopwatch.StartNew();
                                Thread.Sleep(3000);
                                s.Stop();

                            }

                            Thread.Sleep(1000);

                            indataTestVoltage = port.ReadExisting();
                            string tempstr = indataTestVoltage;


                            if (tempstr != null)
                            {
                                if (tempstr.ToLowerInvariant().IndexOf('\n') == -1)
                                {
                                    //  Console.WriteLine("full length didn't come");
                                    var stop = Stopwatch.StartNew();
                                    Thread.Sleep(3000);
                                    stop.Stop();
                                    indataTestVoltage = port.ReadExisting();
                                    indataTestVoltage = tempstr + indataTestVoltage;
                                }
                            }

                            //Console.WriteLine(j + "=" + indataVoltage);
                            temp1++;

                        }
                        if (indataTestVoltage == "" || indataTestVoltage == null)
                            this.excecutionErrorSummery(Int32.Parse(deviceDatalist[j].DeviceId), "Test voltage");
                        TestVoltageCalibrationDatafromRegisters.Add(indataTestVoltage);
                        port.DiscardInBuffer();
                        port.DiscardOutBuffer();
                        port.Close();
                
                     
                    }
                    catch (Exception ex) { Console.WriteLine(ex); continue; }


                }

                TestVoltageCalculations(devicesData, VoltageDataObjects, label11, label12,label8);
            }
            catch (Exception ex)
            {   
                MessageBox.Show(ex.ToString());
            }
        }

        public void TestVoltageCalculations(List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> devicesData, List<voltagelistfromgatherdData> voltagelist, Label label11, Label label12,Label label8)
        {
            try
            {
                var TestVoltageDataObjects = new List<testvoltagelistfromgatherdData>();
                List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;
                GetdataWT1600 getdata = new GetdataWT1600();
                getdata.OpenDevice();
                string[] WT1600data = getdata.writeAndtakeAlltheParametersForVoltage().Split(',');
                getdata.CloseDevice();
                for (int k = 0; k < deviceDatalist.Count; k++)
                {
                    try
                    {
                        testvoltagelistfromgatherdData temp = new testvoltagelistfromgatherdData();
                        temp.DeviceId = deviceDatalist[k].DeviceId;
                        string[] phases2 = TestVoltageCalibrationDatafromRegisters[k].Split(',');
                        temp.testPhaseA = phases2[1];
                        temp.testPhaseB = phases2[2];
                        temp.testPhaseC = phases2[3];
                        //  getdata.CloseDevice();
                        temp.testWt1600ReadVoltage = WT1600data[1];
                        

                        TestVoltageDataObjects.Add(temp);
                    }
                    catch (Exception ex) { Console.WriteLine(ex); continue; }

                }
                Equations eqClass = new Equations();
                eqClass.VoltageCalibarionCalculations(TestVoltageDataObjects, voltagelist);
                label11.Text = " "; label11.Refresh();
                label12.Text = " "; label12.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
    
        //phase

        internal void PhaseDoTheProcessAndReadValues(List<Form1.EnableListDeviceInformation> devicesData,Label label8,Label label11,Label label12)
        {
            List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;

            try
            {
                try
                {
                    GetdataWT1600 getdata3 = new GetdataWT1600();
                    getdata3.OpenDevice();
                    string[] WT1600data2 = getdata3.writeAndtakeAlltheParameters("testvoltage").Split(',');
                    getdata3.CloseDevice();
                    label11.Text = WT1600data2[0];
                    label12.Text = WT1600data2[1];
                    label11.Refresh(); label12.Refresh();
                }
                catch { }
                List<SerialPort> openList = new List<SerialPort>();
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    SerialPort port = new SerialPort(deviceDatalist[i].Port);
                    int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                    int CT = Int32.Parse(deviceDatalist[i].Ct);

                    // Set the properties.
                    port.BaudRate = 115200;
                    port.Parity = Parity.None;
                    // port.ReadTimeout = 1000;
                    port.StopBits = StopBits.One;
                    // //port.RtsEnable = true;
                    // port.Close();
                    // Write a message into the port.
                    openList.Add(port);
                    try
                    {
                        port.Open();
                        var stopwatchcCal = Stopwatch.StartNew();
                        Thread.Sleep(1000);
                        stopwatchcCal.Stop();
                        //write the messages ,can use relevent CT and deviceId
                        //write the messages ,can use relevent CT and deviceId
                        port.ReadExisting();
                        port.Write("Ph_CAL");
                    }
                    catch (Exception ex)
                    {
                        this.excecutionErrorSummery(deviceDatalist[i].Port);
                        deviceDatalist.RemoveAt(i);
                        continue;
                    }
                }
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        string status = "Initial Phase : " + (i + 1) + "/" + deviceDatalist.Count + "(Total devices)";
                        label8.Text = status;
                        label8.Refresh();

                        


                        //do whatever need to do to these ports one by one
                        SerialPort port = openList[i];
                        int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[i].Ct);

                        

                        Thread.Sleep(1000);
                        int temp1 = 0; indataPhase = null;
                        while (indataPhase == null || indataPhase == "")
                        {
                            if (temp1 == NumberOfAttemptsForData) { 
                                 break; }
                            if (temp1 > NumberOfAttemptsForData/2)
                            {
                                port.Write("Ph_CAL");

                                Thread.Sleep(3000);

                            }
                            var stopwatch = Stopwatch.StartNew();
                            Thread.Sleep(1000);
                            stopwatch.Stop();

                            indataPhase = port.ReadExisting();

                            string tempstr = indataTestphase;
                            string s1 = "\n";
                            //  bool b = s1.Contains(indataTestphase);
                            if (tempstr != null)
                            {
                                if (tempstr.ToLowerInvariant().IndexOf('\n') == -1)
                                {
                                    var stop = Stopwatch.StartNew();
                                    Thread.Sleep(3000);
                                    stop.Stop();
                                    indataPhase = port.ReadExisting();
                                    indataPhase = tempstr + indataTestphase;
                                }
                            }
                            temp1++;

                        }
                        if (indataPhase == "" || indataPhase == null)
                            this.excecutionErrorSummery(Int32.Parse(deviceDatalist[i].DeviceId), "Initial phase");
                        PhaseCalibrationDatafromRegisters.Add(indataPhase);
                        port.DiscardInBuffer();
                        port.DiscardOutBuffer();
                        port.Close();
                    }
                    catch (Exception ex) { Console.WriteLine(ex); continue; }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }



        internal void DothePhaseCalculations(List<Form1.EnableListDeviceInformation> devicesData, string mainControllerPort, Label label11, Label label12,Label label8)
        {
            try
            {
                var PhaseDataObjects = new List<PhaselistfromgatherdData>();
                List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;
                GetdataWT1600 getdata = new GetdataWT1600();
                getdata.OpenDevice();
                string[] WT1600data = getdata.writeAndtakeAlltheParameters("phasenormal").Split(',');
                getdata.CloseDevice();
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        PhaselistfromgatherdData temp = new PhaselistfromgatherdData();
                        temp.DeviceId = deviceDatalist[i].DeviceId;
                        string[] phases = PhaseCalibrationDatafromRegisters[i].Split(',');
                        temp.PhaseA = phases[1];
                        temp.PhaseB = phases[2];
                        temp.PhaseC = phases[3];
                        temp.CT = deviceDatalist[i].Ct;
                        temp.port = deviceDatalist[i].Port;
                       
                        temp.Wt1600ReadPhase = WT1600data[2];
                        PhaseDataObjects.Add(temp);
                    }
                    catch (Exception ex) { Console.WriteLine(ex); continue; }

                }

                SerialPort port2 = new SerialPort(mainControllerPort);

                // Set the properties.
                port2.BaudRate = 115200;
                port2.Parity = Parity.None;
                port2.ReadTimeout = 1000;
                port2.StopBits = StopBits.One;

                // Write a message into the port.
                try
                {
                    port2.Open();
                    port2.Write("CON1,187,255,191,1,\n");
                    Thread.Sleep(1000);
                    port2.Write("CON1,187,255,191,0,\n");
                    Thread.Sleep(1000);
                    port2.Write("CON1,125,12,191,0,\n");  //test phase change
                    Thread.Sleep(1000);
                    port2.Write("CON1,125,12,191,1,\n"); //test phase with main
                    Thread.Sleep(1000);
                    port2.Close();
                }
                catch (Exception ex)
                {
                    this.excecutionErrorSummery("Can't connect to main controller!");
                    System.Environment.Exit(1);
                }
             
                Thread.Sleep(6000);
           
                //change the main here,then call testcurrent

                try
                {
                    GetdataWT1600 getdata3 = new GetdataWT1600();
                    getdata3.OpenDevice();
                    string[] WT1600data2 = getdata3.writeAndtakeAlltheParameters("testphase").Split(',');
                    getdata3.CloseDevice();

                    label11.Text = WT1600data2[0];
                    label12.Text = WT1600data2[1];
                    label11.Refresh(); label12.Refresh();
                }
                catch { }
                List<SerialPort> openList = new List<SerialPort>();
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    SerialPort port = new SerialPort(deviceDatalist[i].Port);
                    int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                    int CT = Int32.Parse(deviceDatalist[i].Ct);

                    // Set the properties.
                    port.BaudRate = 115200;
                    port.Parity = Parity.None;
                    // port.ReadTimeout = 1000;
                    port.StopBits = StopBits.One;
                    // //port.RtsEnable = true;
                    // port.Close();
                    // Write a message into the port.
                    openList.Add(port);
                    try
                    {
                        port.Open();
                        var stopwatchcCal = Stopwatch.StartNew();
                        Thread.Sleep(1000);
                        stopwatchcCal.Stop();
                        //write the messages ,can use relevent CT and deviceId
                        //write the messages ,can use relevent CT and deviceId
                        port.ReadExisting();
                        port.Write("Ph_CAL");
                    }
                    catch (Exception ex)
                    {
                        this.excecutionErrorSummery(deviceDatalist[i].Port);
                        deviceDatalist.RemoveAt(i);
                        continue;
                    }
                }
                for (int i = 0; i < deviceDatalist.Count; i++)
                {

                    string status = "Test Phase : " + (i + 1) + "/" + deviceDatalist.Count + "(Total devices)";
                    label8.Text = status;
                    label8.Refresh();

                    
                    //do whatever need to do to these ports one by one
                    SerialPort port = openList[i];
                    int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                    int CT = Int32.Parse(deviceDatalist[i].Ct);

                    
                
                    Thread.Sleep(1000);
                 

                    int temp1 = 0;
                    indataTestphase = null;
                    while (indataTestphase == null || indataTestphase == "")
                    {
                        Console.WriteLine(temp1);
                        if (temp1 == NumberOfAttemptsForData) { 
                        Console.WriteLine("Data not recieving"); break;
                        }
                        if (temp1 > NumberOfAttemptsForData/2)
                        {   
                            port.Write("Ph_CAL");
                            var s = Stopwatch.StartNew();
                            Thread.Sleep(3000);
                            s.Stop();
                        }
                        var stopwatch = Stopwatch.StartNew();
                        Thread.Sleep(1000);
                        stopwatch.Stop();
                        indataTestphase = port.ReadExisting();
                        Console.WriteLine("Recieve data is " + indataTestphase);

                        string tempstr = indataTestphase;
                        string s1 = "\n";
                       // bool b = s1.Contains(indataTestphase);
                        if (tempstr != null)
                        {
                            if (tempstr.ToLowerInvariant().IndexOf('\n') == -1)
                            {
                              
                                Thread.Sleep(3000);
                               
                                indataTestphase = port.ReadExisting();
                                indataTestphase = tempstr + indataTestphase;
                            }
                        }
                        temp1++;

                    }
                    if (indataTestphase == "" || indataTestphase == null)
                        this.excecutionErrorSummery(Int32.Parse(deviceDatalist[i].DeviceId), "Test phase");
                    TestPhaseCalibrationDatafromRegisters.Add(indataTestphase);
                    port.DiscardInBuffer();
                    port.DiscardOutBuffer();
                    port.Close();

           

                }

                var stopwatchcurrent = Stopwatch.StartNew();
                Thread.Sleep(2000);
                stopwatchcurrent.Stop();
                TestPhaseCalculations(devicesData, PhaseDataObjects, label11, label12,label8);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void TestPhaseCalculations(List<Form1.EnableListDeviceInformation> devicesData, List<PhaselistfromgatherdData> PhaseDataObjects, Label label11, Label label12,Label label8)
        {
            try
            {
                var TestPhaseDataObjects = new List<TestPhaselistfromgatherdData>();
                List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;
                GetdataWT1600 getdata = new GetdataWT1600();
                getdata.OpenDevice();
                string[] WT1600data = getdata.writeAndtakeAlltheParameters("testphasenormal").Split(',');
                getdata.CloseDevice();
                for (int k = 0; k < deviceDatalist.Count; k++)
                {
                    try
                    {
                        TestPhaselistfromgatherdData temp = new TestPhaselistfromgatherdData();
                        temp.DeviceId = deviceDatalist[k].DeviceId;
                        string[] phases2 = TestPhaseCalibrationDatafromRegisters[k].Split(',');
                        temp.testPhaseA = phases2[1];
                        temp.testPhaseB = phases2[2];
                        temp.testPhaseC = phases2[3];

                        //get the correct value from wt1600 

                        temp.testWt1600ReadPhase = WT1600data[2];
                        TestPhaseDataObjects.Add(temp);
                    }
                    catch (Exception ex)
                    {
                        continue;
                    }



                }
                Equations eqClass = new Equations();
                eqClass.PhaseCalibarionCalculations(TestPhaseDataObjects, PhaseDataObjects);
                label11.Text = " "; label11.Refresh();
                label12.Text = " "; label12.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

///energy
        ///
        public void ActiveEnergyDoTheProcessAndReadValues(List<Form1.EnableListDeviceInformation> devicesData,string mainport,Label label8,Label label11,Label label12)
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
                try
                {
                    portMain.Open();

                    portMain.Write("CON1,125,28,191,1,\n");
                    Thread.Sleep(1000);
                    portMain.Write("CON1,125,28,191,0,\n");
                    Thread.Sleep(1000);
                    portMain.Write("CON1,127,255,191,0,\n");
                    Thread.Sleep(1000);
                    portMain.Write("CON1,127,255,191,1,\n");
                    Thread.Sleep(1000);
                    portMain.Close();
                }
                catch (Exception ex)
                {
                    this.excecutionErrorSummery("Can't connect to main controller!");
                    System.Environment.Exit(1);
                }
               

              
                Thread.Sleep(6000);
               
               // get the correct value from wt1600

                //for display
                try
                {
                    GetdataWT1600 gDisplay = new GetdataWT1600();
                    gDisplay.OpenDevice();
                    string[] WT1600data3 = gDisplay.writeAndtakeAlltheParameters("firstCurrent").Split(',');
                    gDisplay.CloseDevice();

                    label8.Text = "Active energry LC1 ";
                    label8.Refresh();

                    label11.Text = WT1600data3[0]; label11.Refresh();
                    label12.Text = WT1600data3[1]; label12.Refresh();
                }
                catch { }
                //////for display
           
                Thread.Sleep(3000);
         

                GetdataWT1600Energy getActiveEnergy = new GetdataWT1600Energy();
                getActiveEnergy.OpenDevice();
                string time = "0,1,40";
                getActiveEnergy.startIntegration(time,"high");

                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        label8.Text = "Active energry LC1 " + (i + 1) + "/" + deviceDatalist.Count + "(Total devices)";
                        label8.Refresh();
                        //do whatever need to do to these ports one by one

                        SerialPort port = new SerialPort();
                        port.PortName = deviceDatalist[i].Port;
                        int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[i].Ct);

                        // Set the properties.
                        port.BaudRate = 115200;
                        port.Parity = Parity.None;
                        port.ReadTimeout = 1000;
                        port.StopBits = StopBits.One;
                       // //port.RtsEnable = true;
                        ports.Add(port);
                        // Write a message into the port.
                        try
                        {
                            port.Open();
                            int lineCircles = 10000;
                            linecycle = lineCircles.ToString();
                            port.ReadExisting();
                            string command = "AP_CAL," + lineCircles + ",";
                            port.Write(command);
                            if (deviceDatalist.Count - 1 == i)
                                Thread.Sleep(1000);
                        }
                        catch (Exception ex)
                        {
                            this.excecutionErrorSummery(deviceDatalist[i].Port);
                            deviceDatalist.RemoveAt(i);
                            continue;
                        }
                      
                        //for testing
                        
                    }
                    catch (Exception ex) { Console.WriteLine(ex); continue; }
                    //  port.Close();

                }
                //wait till lineCycle complets and collect data
                //double secondsDouble = Convert.ToDouble("2000");
                //for testing 
                double secondsDouble = Convert.ToDouble("10000");

                int linecircles = (int)(Math.Ceiling(secondsDouble));
                ///wait
                ///
               
                Thread.Sleep(linecircles * 10 + 7000);
         
                //////////////////////////////////////////////////////////////////////////////////////////////
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[i].Ct);
                        int temp1 = 0;
                        while (activeEnergy == null || activeEnergy == "")
                        {
                            Console.WriteLine(temp1);
                            if (temp1 >= NumberOfAttemptsForData) {  break; }
                            var stopwatch = Stopwatch.StartNew();
                            Thread.Sleep(1000);
                            stopwatch.Stop();
                            activeEnergy = ports[i].ReadExisting();
                            Console.WriteLine("recievedData is " + activeEnergy);
                            
                            temp1++;
                        }
                        if (activeEnergy == "" || activeEnergy == null)
                            this.excecutionErrorSummery(Int32.Parse(deviceDatalist[i].DeviceId), "Initial active energy");
                        ActiveEnergyDatafromRegisters.Add(activeEnergy);
                        ports[i].Close(); activeEnergy = null;
                    }
                    catch (Exception ex) { Console.WriteLine(ex); continue; }
                }

                //finally finish the integration on meter and take the value
                string[] WT1600data = getActiveEnergy.returnEnergy("Energynormal").Split(',');
                energyFromTheMeterForActive = WT1600data[0];

                getActiveEnergy.CloseDevice();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                GetdataWT1600Energy getActiveEnergy = new GetdataWT1600Energy();
                getActiveEnergy.CloseDevice();

            }
           
        }

        internal void DotheActiveEnergyCalculations(List<Form1.EnableListDeviceInformation> devicesData, string mainControllerPort,Label label8,Label label11,Label label12)
        {
            try
            {
                var ActiveEnergyObjects = new List<ActiveEnergylistfromgatherdData>();
                List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        ActiveEnergylistfromgatherdData temp = new ActiveEnergylistfromgatherdData();
                        temp.DeviceId = deviceDatalist[i].DeviceId;
                        string[] phases = ActiveEnergyDatafromRegisters[i].Split(',');
                        temp.phaseAenergy = phases[1];
                        temp.phaseBenergy = phases[2];
                        temp.phaseCenergy = phases[3];
                        temp.CT = deviceDatalist[i].Ct;
                        temp.Wt1600MeasuredEnergy = energyFromTheMeterForActive;
                        temp.lineCircle = linecycle;
                        temp.port = deviceDatalist[i].Port;
                        ActiveEnergyObjects.Add(temp);
                        
                    }

                    catch { continue; }
                }

                SerialPort port2 = new SerialPort(mainControllerPort);

                // Set the properties.
                port2.BaudRate = 115200;
                port2.Parity = Parity.None;
                port2.ReadTimeout = 1000;
                port2.StopBits = StopBits.One;

                // Write a message into the port.
                try
                {
                    port2.Open();
                    port2.Write("CON1,127,255,191,1,\n");
                    Thread.Sleep(1000);
                    port2.Write("CON1,127,255,191,0,\n");
                    Thread.Sleep(1000);
                    port2.Write("CON1,159,191,191,0,\n");
                    Thread.Sleep(1000);
                    port2.Write("CON1,159,191,191,1,\n");
                    Thread.Sleep(1000);
                    port2.Write("CON1,159,191,191,1,\n"); //test energy change
                    Thread.Sleep(500);
                    port2.Close();
                }
                catch (Exception ex)
                {
                    this.excecutionErrorSummery("Can't connect to main controller!");
                    System.Environment.Exit(1);
                }
                Thread.Sleep(6000);
                

                //change the main here,start the meter
                label8.Text = "Active energry LC2 ";
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
                        label8.Text = "Active energry LC2 " + (i + 1) + "/" + deviceDatalist.Count;
                        label8.Refresh();
                        //do whatever need to do to these ports one by one
                        SerialPort port = new SerialPort(deviceDatalist[i].Port);
                        int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[i].Ct);

                        // Set the properties.
                        port.BaudRate = 115200;
                        port.Parity = Parity.None;
                        port.ReadTimeout = 1000;
                        port.StopBits = StopBits.One;
                       // //port.RtsEnable = true;
                        // Write a message into the port.
                        try
                        {
                            port.Open();
                            Testports.Add(port);
                            var stopwatchPhCal = Stopwatch.StartNew();
                            Thread.Sleep(1000);
                            stopwatchPhCal.Stop();
                            //write the messages ,can use relevent CT and deviceId
                            int lineCircles = 200;
                            testLineCycle = lineCircles.ToString();
                            port.ReadExisting();
                            string command = "AP_CAL," + lineCircles + ",";
                            port.Write(command);
                            Thread.Sleep(500);
                        }
                        catch (Exception ex)
                        {
                            this.excecutionErrorSummery(deviceDatalist[i].Port);
                            deviceDatalist.RemoveAt(i);
                            continue;
                        }
                    }
                    catch { continue; }
                }

                double secondsDouble = Convert.ToDouble("200");
                
                int linecircles = (int)(Math.Ceiling(secondsDouble));
                ///wait
                var stopwatchenergy = Stopwatch.StartNew();
                Thread.Sleep(linecircles * 10 + 5000);
             
                stopwatchenergy.Stop();

                //////////////////////////////////////////////////////////////////////////////////////////////
                for (int i = 0; i < deviceDatalist.Count; i++)
                {
                    try
                    {
                        //SerialPort port = new SerialPort(deviceDatalist[i].Port);
                        int deviceId = Int32.Parse(deviceDatalist[i].DeviceId);
                        int CT = Int32.Parse(deviceDatalist[i].Ct);

                        int temp1 = 0;
                        while (testActiveEnergy == null || testActiveEnergy == "")
                        {
                            Console.WriteLine(temp1);
                            if (temp1 >= NumberOfAttemptsForData) { break; }
                            var stopwatch = Stopwatch.StartNew();
                            Thread.Sleep(1000);
                            stopwatch.Stop();
                            testActiveEnergy = Testports[i].ReadExisting();
                            Console.WriteLine(testActiveEnergy);
                            temp1++;
                        }
                        if (testActiveEnergy == "" || testActiveEnergy == null)
                            this.excecutionErrorSummery(Int32.Parse(deviceDatalist[i].DeviceId), "Test active energy");
                        TestActiveEnergyCalibrationDatafromRegisters.Add(testActiveEnergy);
                        testActiveEnergy = null;
                        
                    }

                    catch { continue; }

                    //close the main
                 
  
                        try
                        {
                        ports[i].DiscardInBuffer();
                        }
                        catch { }
                        Testports[i].Close();
                }

                string[] WT1600data = getActiveEnergy.returnEnergy("testActive").Split(',');
                testenergyFromTheMeterForActive = WT1600data[0];

                getActiveEnergy.CloseDevice();
                ///////////////////////////////////////////////////////////////////////////////////////

                var stopwatchcurrent = Stopwatch.StartNew();
                Thread.Sleep(2000);
                stopwatchcurrent.Stop();
                TestActiveEnergyCalculations(devicesData, ActiveEnergyObjects);


            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void TestActiveEnergyCalculations(List<Form1.EnableListDeviceInformation> devicesData, List<ActiveEnergylistfromgatherdData> ActiveEnergyDataObjects)
        {
            var TestEnergyDataObjects = new List<TestActiveEnergylistfromgatherdData>();
            List<WindowsFormsApplication5.Form1.EnableListDeviceInformation> deviceDatalist = devicesData;
            for (int k = 0; k < deviceDatalist.Count; k++)
            {
                try
                {
                    TestActiveEnergylistfromgatherdData temp = new TestActiveEnergylistfromgatherdData();
                    temp.DeviceId = deviceDatalist[k].DeviceId;
                    string[] phases2 = TestActiveEnergyCalibrationDatafromRegisters[k].Split(',');
                    temp.TestphaseAenergy = phases2[1];
                    temp.TestphaseBenergy = phases2[2];
                    temp.TestphaseCenergy = phases2[3];

                    //get the correct value from wt1600
                    temp.TestWt1600MeasuredEnergy = testenergyFromTheMeterForActive;
                    temp.TestlineCircle = testLineCycle;

                    TestEnergyDataObjects.Add(temp);
                }
                catch (Exception ex) { Console.WriteLine(ex); continue; }


            }
            Equations eqClass = new Equations();
            eqClass.EneryCalibarionCalculations(TestEnergyDataObjects, ActiveEnergyDataObjects);
        }

//Reactive Energy

        



        public class currentlistfromgatherdData
        {
            public string CT { get; set; }
            public string port { get; set; }
            public string DeviceId { get; set; }
            public string Wt1600Readcurrent { get; set; }
            public string PhaseA { get; set; }
            public string PhaseB { get; set; }
            public string PhaseC { get; set; }
        }

        public class testcurrentfromgatherdata 
        {
            public string DeviceId { get; set; }
            public string testWt1600Readcurrent { get; set; }
            public string testPhaseA { get; set; }
            public string testPhaseB { get; set; }
            public string testPhaseC { get; set; }
        }

        public class voltagelistfromgatherdData
        {
            public string DeviceId { get; set; }
            public string port { get; set; }
            public string Wt1600ReadVoltage { get; set; }
            public string PhaseA { get; set; }
            public string PhaseB { get; set; }
            public string PhaseC { get; set; }
        }

        public class testvoltagelistfromgatherdData
        {
            public string DeviceId { get; set; }
            public string testWt1600ReadVoltage { get; set; }
            public string testPhaseA { get; set; }
            public string testPhaseB { get; set; }
            public string testPhaseC { get; set; }
        }

        public class PhaselistfromgatherdData
        {
            public string CT { get; set; }
            public string DeviceId { get; set; }
            public string Wt1600ReadPhase { get; set; }
            public string PhaseA { get; set; }
            public string PhaseB { get; set; }
            public string PhaseC { get; set; }
            public string port { get; set; }
        }

        public class TestPhaselistfromgatherdData
        {
            public string usedpowerfactor { get; set; }
            public string DeviceId { get; set; }
            public string testWt1600ReadPhase { get; set; }
            public string testPhaseA { get; set; }
            public string testPhaseB { get; set; }
            public string testPhaseC { get; set; }
        }

        public class ActiveEnergylistfromgatherdData
        {
            public string CT { get; set; }
            public string DeviceId { get; set; }
            public string Wt1600MeasuredEnergy { get; set; }
            public string phaseAenergy { get; set; }
            public string phaseBenergy { get; set; }
            public string phaseCenergy { get; set; }
            public string lineCircle { get; set; }
            public string port { get; set; }

        }

        public class TestActiveEnergylistfromgatherdData
        {
            public string DeviceId { get; set; }
            public string TestWt1600MeasuredEnergy { get; set; }
            public string TestphaseAenergy { get; set; }
            public string TestphaseBenergy { get; set; }
            public string TestphaseCenergy { get; set; }
            public string TestlineCircle { get; set; }
        }

        public static void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            string indata = sp.ReadExisting();
            Console.WriteLine("Data Received: ");
            Console.Write(indata);


            //Call to the splitting function here
            string GotString = indata;
            currentPortData = indata;
            // Console.ReadLine();
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
                sw.WriteLine("Field :"+field+" from device :"+id+" was not recieved!");
            }	
        }
        public void excecutionErrorSummery(string id)
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
                sw.WriteLine("Port :" + id + " is disconnected!");
            }
        }
    }
    
}



