using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace WindowsFormsApplication5
{
    class Equations
    {

        internal void currentCalibarionCalculations(List<ExecutionOfProgramme.testcurrentfromgatherdata> TestCurrentDataObjects, List<ExecutionOfProgramme.currentlistfromgatherdData> CurrentDataObjects)
        {
            try
            {
                List<ExecutionOfProgramme.currentlistfromgatherdData> CurrentcaliDataObjectlist = CurrentDataObjects;
                List<ExecutionOfProgramme.testcurrentfromgatherdata> TestCurrentcaliDataObjectlist = TestCurrentDataObjects;
                //  var finalcurrentObjects = new List<currentlistTowriteFile>();
                ///////////////////for each device do the following and write to a csv file
                for (int i = 0; i < CurrentcaliDataObjectlist.Count; i++)
                {
                    try
                    {
                        int facter;
                        switch (CurrentcaliDataObjectlist[i].CT)
                        {
                            case "30":
                                facter = 1;
                                break;
                            case "60":
                                facter = 2;
                                break;
                            case "100":
                                facter = 3;
                                break;
                            case "200":
                                facter = 6;
                                break;
                            case "400":
                                facter = 13;
                                break;
                            case "600":
                                facter = 20;
                                break;
                            case "1000":
                                facter = 33;
                                break;
                            case "1200":
                                facter = 40;
                                break;
                            case "1500":
                                facter = 50;
                                break;
                            default:
                                facter = 1;
                                break;
                        }

                        currentlistTowriteFile temp = new currentlistTowriteFile();
                        decimal fulscaleCurrentdiv500 = decimal.Parse(CurrentcaliDataObjectlist[i].Wt1600Readcurrent) * facter;
                        decimal PhaseARegValuForFcurrentDiv50 = decimal.Parse(CurrentcaliDataObjectlist[i].PhaseA);
                        decimal PhaseBRegValuForFcurrentDiv50 = decimal.Parse(CurrentcaliDataObjectlist[i].PhaseB);
                        decimal PhaseCRegValuForFcurrentDiv50 = decimal.Parse(CurrentcaliDataObjectlist[i].PhaseC);
                        decimal testCurrent = decimal.Parse(TestCurrentcaliDataObjectlist[i].testWt1600Readcurrent) * facter;
                        decimal RegvalueFORTestCurrent = decimal.Parse(TestCurrentcaliDataObjectlist[i].testPhaseA);
                        decimal RegvaluePhaseBFORTestCurrent = decimal.Parse(TestCurrentcaliDataObjectlist[i].testPhaseB);
                        decimal RegvaluePhaseCFORTestCurrent = decimal.Parse(TestCurrentcaliDataObjectlist[i].testPhaseC);
                        // decimal RegvalueFORTestCurrent = decimal.Parse(TestCurrentcaliDataObjectlist[i].testPhaseA);

                        decimal IRMSOS = (testCurrent * testCurrent * PhaseARegValuForFcurrentDiv50 * PhaseARegValuForFcurrentDiv50 - ((fulscaleCurrentdiv500 * fulscaleCurrentdiv500) * (RegvalueFORTestCurrent * RegvalueFORTestCurrent))) / (16384 * ((fulscaleCurrentdiv500 * fulscaleCurrentdiv500) - (testCurrent * testCurrent)));
                        decimal IRMSOSPhaseB = (testCurrent * testCurrent * PhaseBRegValuForFcurrentDiv50 * PhaseARegValuForFcurrentDiv50 - ((fulscaleCurrentdiv500 * fulscaleCurrentdiv500) * (RegvaluePhaseBFORTestCurrent * RegvaluePhaseBFORTestCurrent))) / (16384 * ((fulscaleCurrentdiv500 * fulscaleCurrentdiv500) - (testCurrent * testCurrent)));
                        decimal IRMSOSPhaseC = (testCurrent * testCurrent * PhaseCRegValuForFcurrentDiv50 * PhaseARegValuForFcurrentDiv50 - ((fulscaleCurrentdiv500 * fulscaleCurrentdiv500) * (RegvaluePhaseCFORTestCurrent * RegvaluePhaseCFORTestCurrent))) / (16384 * ((fulscaleCurrentdiv500 * fulscaleCurrentdiv500) - (testCurrent * testCurrent)));
                        //  String  hexValue_currentcalibraion = IRMSOS.ToString("X");
                        temp.deviceID = CurrentcaliDataObjectlist[i].DeviceId;
                        temp.fulscaleCurrentdiv500 = fulscaleCurrentdiv500;
                        temp.PhaseARegValuForFcurrentDiv50 = PhaseARegValuForFcurrentDiv50;
                        temp.PhaseBRegValuForFcurrentDiv50 = PhaseBRegValuForFcurrentDiv50;
                        temp.PhaseCRegValuForFcurrentDiv50 = PhaseCRegValuForFcurrentDiv50;
                        temp.testCurrent = testCurrent;
                        temp.RegvaluePhaseAforTestCurrent = RegvalueFORTestCurrent;
                        temp.RegvaluePhaseBforTestCurrent = RegvaluePhaseBFORTestCurrent;
                        temp.RegvaluePhaseCforTestCurrent = RegvaluePhaseCFORTestCurrent;
                        temp.IRMSOSPhaseA = IRMSOS;
                        temp.IRMSOSPhaseB = IRMSOSPhaseB;
                        temp.IRMSOSPhaseC = IRMSOSPhaseC;
                        temp.port = CurrentcaliDataObjectlist[i].port;
                        filewriter(temp,CurrentcaliDataObjectlist[i].CT);
//write observed offset parameters to the epro
                        SerialPort port = new SerialPort();
                        port.PortName = CurrentcaliDataObjectlist[i].port;
                        port.BaudRate = 115200;
                        port.Parity = Parity.None;
                        port.StopBits = StopBits.One;
                        port.RtsEnable = true;
                        try
                        {
                            port.Open();
                            String portTemporyOffsets = "OFFSET_C," + Math.Round(IRMSOS) + "," + Math.Round(IRMSOSPhaseB) + "," + Math.Round(IRMSOSPhaseC) + ",";
                            port.DiscardOutBuffer();

                            port.Write(portTemporyOffsets);
                            Thread.Sleep(50);
                            port.DiscardInBuffer();
                            port.DiscardOutBuffer();
                            port.Close();
                        }
                        catch (Exception ex)
                        {
                            this.excecutionErrorSummery(temp.deviceID, "current");
                        }
                        
                      ////////////////////////////
                    }
                    catch { continue; }
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            /////////////////////

        }

        internal void VoltageCalibarionCalculations(List<ExecutionOfProgramme.testvoltagelistfromgatherdData> TestVoltageDataObjects, List<ExecutionOfProgramme.voltagelistfromgatherdData> voltagelist)
        {
            try
            {
                List<ExecutionOfProgramme.voltagelistfromgatherdData> VoltagecaliDataObjectlist = voltagelist;
                List<ExecutionOfProgramme.testvoltagelistfromgatherdData> testvoltagelistfromgatherdData = TestVoltageDataObjects;
                for (int i = 0; i < voltagelist.Count; i++)
                {
                    try
                    {
                        VoltagelistTowriteFile temp = new VoltagelistTowriteFile();
                        decimal fulscaleVoltageDiv20 = decimal.Parse(voltagelist[i].Wt1600ReadVoltage);
                        decimal PhaseARegValuForFVoltageDiv20 = decimal.Parse(voltagelist[i].PhaseA);
                        decimal PhaseBRegValuForFVoltageDiv20 = decimal.Parse(voltagelist[i].PhaseB);
                        decimal PhaseCRegValuForFVoltageDiv20 = decimal.Parse(voltagelist[i].PhaseC);


                        decimal testVoltage = decimal.Parse(TestVoltageDataObjects[i].testWt1600ReadVoltage);
                        decimal testRegvalForVoltagePhaseA = decimal.Parse(testvoltagelistfromgatherdData[i].testPhaseA);
                        decimal testRegvalForVoltagePhaseB = decimal.Parse(testvoltagelistfromgatherdData[i].testPhaseB);
                        decimal testRegvalForVoltagePhaseC = decimal.Parse(testvoltagelistfromgatherdData[i].testPhaseC);

                        decimal VRMSOS = ((testVoltage * PhaseARegValuForFVoltageDiv20) - (fulscaleVoltageDiv20 * testRegvalForVoltagePhaseA)) / (64 * (fulscaleVoltageDiv20 - testVoltage));
                        decimal PhaseBVRMSOS = ((testVoltage * PhaseBRegValuForFVoltageDiv20) - (fulscaleVoltageDiv20 * testRegvalForVoltagePhaseB)) / (64 * (fulscaleVoltageDiv20 - testVoltage));
                        decimal PhaseCVRMSOS = ((testVoltage * PhaseCRegValuForFVoltageDiv20) - (fulscaleVoltageDiv20 * testRegvalForVoltagePhaseC)) / (64 * (fulscaleVoltageDiv20 - testVoltage));

                        temp.deviceID = voltagelist[i].DeviceId;
                        temp.fulscaleVoltageDiv20 = fulscaleVoltageDiv20;
                        temp.PhaseARegValuForFVoltageDiv20 = PhaseARegValuForFVoltageDiv20;
                        temp.PhaseBRegValuForFVoltageDiv20 = PhaseBRegValuForFVoltageDiv20;
                        temp.PhaseCRegValuForFVoltageDiv20 = PhaseCRegValuForFVoltageDiv20;
                        temp.testVoltage = testVoltage;
                        temp.testRegvalForVoltagePhaseA = testRegvalForVoltagePhaseA;
                        temp.testRegvalForVoltagePhaseB = testRegvalForVoltagePhaseB;
                        temp.testRegvalForVoltagePhaseC = testRegvalForVoltagePhaseC;
                        temp.PhaseAVRMSOS = VRMSOS;
                        temp.PhaseBVRMSOS = PhaseBVRMSOS;
                        temp.PhaseCVRMSOS = PhaseCVRMSOS;
                        // String  hexValue_votagecalibraion = VRMSOS.ToString("X");
                        filewriterforVoltage(temp);
                        //write observed offset parameters to the epro
                        SerialPort port = new SerialPort();
                        port.PortName = VoltagecaliDataObjectlist[i].port;
                        port.BaudRate = 115200;
                        port.Parity = Parity.None;
                        port.StopBits = StopBits.One;
                        port.RtsEnable = true;
                        try
                        {
                            //write to the port    
                            port.Open();


                            String portTemporyOffsets = "OFFSET_V," + Math.Round(VRMSOS) + "," + Math.Round(PhaseBVRMSOS) + "," + Math.Round(PhaseCVRMSOS) + ",";
                            //       MessageBox.Show(portTemporyOffsets);
                            port.DiscardOutBuffer();
                            Thread.Sleep(50);
                            port.Write(portTemporyOffsets);
                            Thread.Sleep(50);
                            port.DiscardInBuffer();
                            port.DiscardOutBuffer();
                            port.Close();
                        }
                        catch (Exception ex)
                        {
                            this.excecutionErrorSummery(temp.deviceID, "voltage");
                        }
                        
                    }
                    catch { continue; }
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        internal void PhaseCalibarionCalculations(List<ExecutionOfProgramme.TestPhaselistfromgatherdData> TestPhaseDataObjects, List<ExecutionOfProgramme.PhaselistfromgatherdData> PhaseDataObjects)
        {
            try
            {
                List<ExecutionOfProgramme.PhaselistfromgatherdData> PhasecaliDataObjectlist = PhaseDataObjects;
                List<ExecutionOfProgramme.TestPhaselistfromgatherdData> testPhaselistfromgatherdData = TestPhaseDataObjects;
                for (int i = 0; i < PhasecaliDataObjectlist.Count; i++)
                {
                    try
                    {
                        PhaselistTowriteFile temp = new PhaselistTowriteFile();

                        decimal Power_at_pf_1_at_Itest_and_Vnom_forPhaseA = decimal.Parse(PhaseDataObjects[i].PhaseA);
                        decimal Power_at_pf_1_at_Itest_and_Vnom_forPhaseB = decimal.Parse(PhaseDataObjects[i].PhaseB);
                        decimal Power_at_pf_1_at_Itest_and_Vnom_forPhaseC = decimal.Parse(PhaseDataObjects[i].PhaseC);

                        decimal Power_at_pf_x_at_Itest_and_VnomforPhaseA = decimal.Parse(TestPhaseDataObjects[i].testPhaseA);
                        decimal Power_at_pf_x_at_Itest_and_VnomforPhaseB = decimal.Parse(TestPhaseDataObjects[i].testPhaseB);
                        decimal Power_at_pf_x_at_Itest_and_VnomforPhaseC = decimal.Parse(TestPhaseDataObjects[i].testPhaseC);

                        double Used_Power_factor = (double)(decimal.Parse(TestPhaseDataObjects[i].testWt1600ReadPhase));

                        double errorPhaseA = ((double)Power_at_pf_x_at_Itest_and_VnomforPhaseA - ((double)Power_at_pf_1_at_Itest_and_Vnom_forPhaseA * Used_Power_factor)) / (double)((double)Power_at_pf_1_at_Itest_and_Vnom_forPhaseA * Used_Power_factor);
                        double errorPhaseB = ((double)Power_at_pf_x_at_Itest_and_VnomforPhaseB - ((double)Power_at_pf_1_at_Itest_and_Vnom_forPhaseB * Used_Power_factor)) / (double)((double)Power_at_pf_1_at_Itest_and_Vnom_forPhaseB * Used_Power_factor);
                        double errorPhaseC = ((double)Power_at_pf_x_at_Itest_and_VnomforPhaseC - ((double)Power_at_pf_1_at_Itest_and_Vnom_forPhaseC * Used_Power_factor)) / (double)((double)Power_at_pf_1_at_Itest_and_Vnom_forPhaseC * Used_Power_factor);

                        double tempA = (double)((decimal)errorPhaseA / Convert.ToDecimal(Math.Sqrt(3)));
                        double tempB = (double)((decimal)errorPhaseB / Convert.ToDecimal(Math.Sqrt(3)));
                        double tempC = (double)((decimal)errorPhaseC / Convert.ToDecimal(Math.Sqrt(3)));

                        double phaseErrorA = -Math.Asin(tempA);
                        double phaseErrorB = -Math.Asin(tempB);
                        double phaseErrorC = -Math.Asin(tempC);

                        double varA = 0; double varB = 0; double varC = 0;
                        if (errorPhaseA < 0) { varA = 1.2; }
                        if (errorPhaseA >= 0) { varA = 2.4; }
                        if (errorPhaseB < 0) { varB = 1.2; }
                        if (errorPhaseB >= 0) { varB = 2.4; }
                        if (errorPhaseC < 0) { varC = 1.2; }
                        if (errorPhaseC >= 0) { varC = 2.4; }
                        decimal xphcalA;decimal xphcalB;decimal xphcalC;
                        if ((phaseErrorA / (varA * 1 * 2 * Math.PI * Math.Pow(10, -6))) <= -64) {  xphcalA = -64; }
                        if ((phaseErrorA / (varA * 1 * 2 * Math.PI * Math.Pow(10, -6))) > 63) { xphcalA = 63; }
                        else { xphcalA=(decimal)(phaseErrorA / (varA * 1 * 2 * Math.PI * Math.Pow(10, -6))); }

                        if ((phaseErrorB / (varB * 1 * 2 * Math.PI * Math.Pow(10, -6))) <= -64) {xphcalB = -64; }
                        if ((phaseErrorB / (varB * 1 * 2 * Math.PI * Math.Pow(10, -6))) > 63) {  xphcalB= 63; }
                        else { xphcalB = (decimal)(phaseErrorB / (varB * 1 * 2 * Math.PI * Math.Pow(10, -6))); }

                        if ((phaseErrorC / (varC * 1 * 2 * Math.PI * Math.Pow(10, -6))) <= -64) { xphcalC = -64; }
                        if ((phaseErrorC / (varC * 1 * 2 * Math.PI * Math.Pow(10, -6))) > 63) {  xphcalC= 63; }
                        else { xphcalC = (decimal)(phaseErrorA / (varC * 1 * 2 * Math.PI * Math.Pow(10, -6))); }
                        
                        temp.deviceID = PhaseDataObjects[i].DeviceId;
                        temp.Power_at_pf_1_at_Itest_and_Vnom_phaseA = Power_at_pf_1_at_Itest_and_Vnom_forPhaseA;
                        temp.Power_at_pf_1_at_Itest_and_Vnom_phaseB = Power_at_pf_1_at_Itest_and_Vnom_forPhaseB;
                        temp.Power_at_pf_1_at_Itest_and_Vnom_phaseC = Power_at_pf_1_at_Itest_and_Vnom_forPhaseC;
                        temp.Power_at_pf_x_at_Itest_and_Vnom_phaseA = Power_at_pf_x_at_Itest_and_VnomforPhaseA;
                        temp.Power_at_pf_x_at_Itest_and_Vnom_phaseB = Power_at_pf_x_at_Itest_and_VnomforPhaseB;
                        temp.Power_at_pf_x_at_Itest_and_Vnom_phaseC = Power_at_pf_x_at_Itest_and_VnomforPhaseC;
                        temp.ErrorA = errorPhaseA;
                        temp.ErrorB = errorPhaseB;
                        temp.ErrorC = errorPhaseC;
                        temp.Phase_ErrorA = (decimal)phaseErrorA;
                        temp.Phase_ErrorB = (decimal)phaseErrorB;
                        temp.Phase_ErrorC = (decimal)phaseErrorC;
                       
                        temp.xPHCALA = xphcalA;
                        temp.xPHCALB = xphcalB;
                        temp.xPHCALC = xphcalC;

                        temp.Used_Power_factor = Used_Power_factor;

                        Console.Write("Phase data requset to writte from Device ID :" + temp.deviceID);
                        filewriterforPhase(temp);
                        Console.Write("Phase data written from Device ID :" + temp.deviceID);

                        SerialPort port = new SerialPort();
                        port.PortName = PhaseDataObjects[i].port;
                        port.BaudRate = 115200;
                        port.Parity = Parity.None;
                        port.StopBits = StopBits.One;
                        port.RtsEnable = true;
                        try
                        {
                            //write to the port    
                            port.Open();
                            String portTemporyOffsets = "OFFSET_Ph," + Math.Round(xphcalA) + "," + Math.Round(xphcalB) + "," + Math.Round(xphcalC) + ",";
                            port.DiscardOutBuffer();
                            Thread.Sleep(50);
                            port.Write(portTemporyOffsets);
                            Thread.Sleep(50);
                            //   MessageBox.Show(portTemporyOffsets);
                            port.DiscardInBuffer();
                            port.DiscardOutBuffer();
                            port.Close();
                        }
                        catch (Exception ex)
                        {
                            this.excecutionErrorSummery(temp.deviceID, "phase");
                        }
                        
                       
                    }
                    catch { continue; }
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            
        }

        internal void EneryCalibarionCalculations(List<ExecutionOfProgramme.TestActiveEnergylistfromgatherdData> TestEnergyDataObjects, List<ExecutionOfProgramme.ActiveEnergylistfromgatherdData> ActiveEnergyDataObjects)
        {
            List<ExecutionOfProgramme.ActiveEnergylistfromgatherdData> EnergycaliDataObjectlist = ActiveEnergyDataObjects;
            List<ExecutionOfProgramme.TestActiveEnergylistfromgatherdData> testEnergylistfromgatherdData = TestEnergyDataObjects;
            for (int i = 0; i < EnergycaliDataObjectlist.Count; i++)
            {
                int facter;
                switch (EnergycaliDataObjectlist[i].CT)
                {
                    case "30":
                        facter = 1;
                        break;
                    case "60":
                        facter = 2;
                        break;
                    case "100":
                        facter = 3;
                        break;
                    case "200":
                        facter = 6;
                        break;
                    case "400":
                        facter = 13;
                        break;
                    case "600":
                        facter = 20;
                        break;
                    case "1000":
                        facter = 33;
                        break;
                    case "1200":
                        facter = 40;
                        break;
                    case "1500":
                        facter = 50;
                        break;
                    default:
                        facter = 1;
                        break;
                }

                try
                {
                    // double Fullscalecurrentdiv500 ;
                    double PhaseACorrespondingWATTHRregValueForFirst = double.Parse(EnergycaliDataObjectlist[i].phaseAenergy);
                    double PhaseBCorrespondingWATTHRregValueForFirst = double.Parse(EnergycaliDataObjectlist[i].phaseBenergy);
                    double PhaseCCorrespondingWATTHRregValueForFirst = double.Parse(EnergycaliDataObjectlist[i].phaseCenergy);
                    double LineCycle1 = double.Parse(EnergycaliDataObjectlist[i].lineCircle);
                    double firstMeasuredEnergy =facter* double.Parse(EnergycaliDataObjectlist[i].Wt1600MeasuredEnergy);
                    //double testCurrent = double.Parse(EnergycaliDataObjectlist[i].t);
                    double TestPhaseACorrespondingWATTHRregValue = double.Parse(TestEnergyDataObjects[i].TestphaseAenergy);
                    double TestPhaseBCorrespondingWATTHRregValue = double.Parse(TestEnergyDataObjects[i].TestphaseBenergy);
                    double TestPhaseCCorrespondingWATTHRregValue = double.Parse(TestEnergyDataObjects[i].TestphaseCenergy);
                    double TestLineCycle = double.Parse(testEnergylistfromgatherdData[i].TestlineCircle);
                    double secondMeasuredEnergy =facter* double.Parse(testEnergylistfromgatherdData[i].TestWt1600MeasuredEnergy);
                    //  double offset = ((CorrespondingWATTHRregValue * testCurrent) - (((TestCorrespondingWATTHRregValue * LineCycle1) / TestLineCycle) * Fullscalecurrentdiv500)) / (Fullscalecurrentdiv500 - testCurrent);

                    double offsetA = ((PhaseACorrespondingWATTHRregValueForFirst * secondMeasuredEnergy) - (TestPhaseACorrespondingWATTHRregValue * firstMeasuredEnergy)) / (((LineCycle1 / TestLineCycle) * firstMeasuredEnergy) - secondMeasuredEnergy);
                    double offsetB = ((PhaseBCorrespondingWATTHRregValueForFirst * secondMeasuredEnergy) - (TestPhaseBCorrespondingWATTHRregValue * firstMeasuredEnergy)) / (((LineCycle1 / TestLineCycle) * firstMeasuredEnergy) - secondMeasuredEnergy);
                    double offsetC = ((PhaseCCorrespondingWATTHRregValueForFirst * secondMeasuredEnergy) - (TestPhaseCCorrespondingWATTHRregValue * firstMeasuredEnergy)) / (((LineCycle1 / TestLineCycle) * firstMeasuredEnergy) - secondMeasuredEnergy);

                    double xWATTOSA = ((offsetA * 4) * (2 * Math.Pow(2, 29))) / (100 * Math.Pow(10, 7));
                    double xWATTOSB = ((offsetB * 4) * (2 * Math.Pow(2, 29))) / (100 * Math.Pow(10, 7));
                    double xWATTOSC = ((offsetC * 4) * (2 * Math.Pow(2, 29))) / (100 * Math.Pow(10, 7));


                    energylistWriteFile temp = new energylistWriteFile();
                    temp.deviceId = EnergycaliDataObjectlist[i].DeviceId;
                    temp.PhaseACorrespondingWATTHRregValueForFirst = PhaseACorrespondingWATTHRregValueForFirst.ToString();
                    temp.PhaseBCorrespondingWATTHRregValueForFirst = PhaseBCorrespondingWATTHRregValueForFirst.ToString();
                    temp.PhaseCCorrespondingWATTHRregValueForFirst = PhaseCCorrespondingWATTHRregValueForFirst.ToString();
                    temp.TestPhaseACorrespondingWATTHRregValue = TestPhaseACorrespondingWATTHRregValue.ToString();
                    temp.TestPhaseBCorrespondingWATTHRregValue = TestPhaseBCorrespondingWATTHRregValue.ToString();
                    temp.TestPhaseCCorrespondingWATTHRregValue = TestPhaseCCorrespondingWATTHRregValue.ToString();
                    temp.LineCycle1 = LineCycle1.ToString();
                    temp.TestLineCycle = TestLineCycle.ToString();
                    temp.energfrimWt16001st = firstMeasuredEnergy.ToString();
                    temp.energfrimWt16002nd = secondMeasuredEnergy.ToString();
                    temp.OffsetA = offsetA.ToString();
                    temp.OffsetB = offsetB.ToString();
                    temp.OffsetC = offsetC.ToString();
                    temp.xWATTOSA = xWATTOSA.ToString();
                    temp.xWATTOSB = xWATTOSB.ToString();
                    temp.xWATTOSC = xWATTOSC.ToString();

                    fileWriterForActiveEnergy(temp);

                    SerialPort port = new SerialPort();
                    port.PortName = EnergycaliDataObjectlist[i].port;
                    port.BaudRate = 115200;
                    port.Parity = Parity.None;
                    port.StopBits = StopBits.One;
                    port.RtsEnable = true;

                    //write to the port    
                    try{
                    port.Open();
                    String portTemporyOffsets = "OFFSET_AP," + Math.Round((xWATTOSA)) + "," + Math.Round(xWATTOSB) + "," + Math.Round(xWATTOSC) + ",";
                    port.DiscardOutBuffer();
                    port.Write(portTemporyOffsets);
                    Thread.Sleep(50);
                    port.DiscardInBuffer();
                    port.DiscardOutBuffer();
                    port.Close();
                    }
                    catch (Exception ex)
                    {
                        this.excecutionErrorSummery(temp.deviceId, "active energy");
                    }
                        
                }
                catch { continue; }
            }

            }

        internal void ReactiveEneryCalibarionCalculations(List<ReactiveEnergyExecution.TestReactiveEnergylistfromgatherdData> TestEnergyDataObjects, List<ReactiveEnergyExecution.ReactiveEnergylistfromgatherdData> ReactiveEnergyDataObjects)
        {
            List<ReactiveEnergyExecution.ReactiveEnergylistfromgatherdData> EnergycaliDataObjectlist = ReactiveEnergyDataObjects;
            List<ReactiveEnergyExecution.TestReactiveEnergylistfromgatherdData> testEnergylistfromgatherdData = TestEnergyDataObjects;
            for (int i = 0; i < EnergycaliDataObjectlist.Count; i++)
            {
                int facter;
                switch (EnergycaliDataObjectlist[i].CT)
                {
                    case "30":
                        facter = 1;
                        break;
                    case "60":
                        facter = 2;
                        break;
                    case "100":
                        facter = 3;
                        break;
                    case "200":
                        facter = 6;
                        break;
                    case "400":
                        facter = 13;
                        break;
                    case "600":
                        facter = 20;
                        break;
                    case "1000":
                        facter = 33;
                        break;
                    case "1200":
                        facter = 40;
                        break;
                    case "1500":
                        facter = 50;
                        break;
                    default:
                        facter = 1;
                        break;
                }
                try
                {
                    // double Fullscalecurrentdiv500 ;
                    double PhaseACorrespondingWATTHRregValueForFirst = double.Parse(EnergycaliDataObjectlist[i].phaseAenergy);
                    double PhaseBCorrespondingWATTHRregValueForFirst = double.Parse(EnergycaliDataObjectlist[i].phaseBenergy);
                    double PhaseCCorrespondingWATTHRregValueForFirst = double.Parse(EnergycaliDataObjectlist[i].phaseCenergy);
                    double LineCycle1 = double.Parse(EnergycaliDataObjectlist[i].lineCircle);
                    double firstMeasuredEnergy =facter* double.Parse(EnergycaliDataObjectlist[i].Wt1600MeasuredReactiveEnergy);
                    //double testCurrent = double.Parse(EnergycaliDataObjectlist[i].t);
                    double TestPhaseACorrespondingWATTHRregValue = double.Parse(TestEnergyDataObjects[i].TestphaseAenergy);
                    double TestPhaseBCorrespondingWATTHRregValue = double.Parse(TestEnergyDataObjects[i].TestphaseBenergy);
                    double TestPhaseCCorrespondingWATTHRregValue = double.Parse(TestEnergyDataObjects[i].TestphaseCenergy);
                    double TestLineCycle = double.Parse(testEnergylistfromgatherdData[i].TestlineCircle);
                    double secondMeasuredEnergy = facter*double.Parse(testEnergylistfromgatherdData[i].TestWt1600MeasuredEnergy);
                    //  double offset = ((CorrespondingWATTHRregValue * testCurrent) - (((TestCorrespondingWATTHRregValue * LineCycle1) / TestLineCycle) * Fullscalecurrentdiv500)) / (Fullscalecurrentdiv500 - testCurrent);
                    double offsetA = ((PhaseACorrespondingWATTHRregValueForFirst * secondMeasuredEnergy) - (TestPhaseACorrespondingWATTHRregValue * firstMeasuredEnergy)) / (((LineCycle1 / TestLineCycle) * firstMeasuredEnergy) - secondMeasuredEnergy);
                    double offsetB = ((PhaseBCorrespondingWATTHRregValueForFirst * secondMeasuredEnergy) - (TestPhaseBCorrespondingWATTHRregValue * firstMeasuredEnergy)) / (((LineCycle1 / TestLineCycle) * firstMeasuredEnergy) - secondMeasuredEnergy);
                    double offsetC = ((PhaseCCorrespondingWATTHRregValueForFirst * secondMeasuredEnergy) - (TestPhaseBCorrespondingWATTHRregValue * firstMeasuredEnergy)) / (((LineCycle1 / TestLineCycle) * firstMeasuredEnergy) - secondMeasuredEnergy);

                    double xWATTOSA = ((offsetA * 4) * (2 * Math.Pow(2, 29))) / (100 * Math.Pow(10, 7));
                    double xWATTOSB = ((offsetB * 4) * (2 * Math.Pow(2, 29))) / (100 * Math.Pow(10, 7));
                    double xWATTOSC = ((offsetC * 4) * (2 * Math.Pow(2, 29))) / (100 * Math.Pow(10, 7));

                    ReactiveEnergylistWriteFile temp = new ReactiveEnergylistWriteFile();
                    temp.deviceId = EnergycaliDataObjectlist[i].DeviceId;
                    temp.PhaseACorrespondingWATTHRregValueForFirst = PhaseACorrespondingWATTHRregValueForFirst.ToString();
                    temp.PhaseBCorrespondingWATTHRregValueForFirst = PhaseBCorrespondingWATTHRregValueForFirst.ToString();
                    temp.PhaseCCorrespondingWATTHRregValueForFirst = PhaseCCorrespondingWATTHRregValueForFirst.ToString();
                    temp.TestPhaseACorrespondingWATTHRregValue = TestPhaseACorrespondingWATTHRregValue.ToString();
                    temp.TestPhaseBCorrespondingWATTHRregValue = TestPhaseBCorrespondingWATTHRregValue.ToString();
                    temp.TestPhaseCCorrespondingWATTHRregValue = TestPhaseCCorrespondingWATTHRregValue.ToString();
                    temp.LineCycle1 = LineCycle1.ToString();
                    temp.TestLineCycle = TestLineCycle.ToString();
                    temp.energfrimWt16001st = firstMeasuredEnergy.ToString();
                    temp.energfrimWt16002nd = secondMeasuredEnergy.ToString();
                    temp.OffsetA = offsetA.ToString();
                    temp.OffsetB = offsetB.ToString();
                    temp.OffsetC = offsetC.ToString();
                    temp.xWATTOSA = xWATTOSA.ToString();
                    temp.xWATTOSB = xWATTOSB.ToString();
                    temp.xWATTOSC = xWATTOSC.ToString();
                    fileWriterForAReactiveEnergy(temp);

                    SerialPort port = new SerialPort();
                    port.PortName = EnergycaliDataObjectlist[i].port;
                    port.BaudRate = 115200;
                    port.Parity = Parity.None;
                    port.StopBits = StopBits.One;
                    port.RtsEnable = true;

                    //write to the port   
                    try
                    {
                        port.Open();
                        String portTemporyOffsets = "OFFSET_RP," + Math.Round((xWATTOSA)) + "," + Math.Round(xWATTOSB) + "," + Math.Round(xWATTOSC) + ",";
                        port.DiscardOutBuffer();
                        port.Write(portTemporyOffsets);
                        Thread.Sleep(50);
                        port.DiscardInBuffer();
                        port.DiscardOutBuffer();
                        port.Close();
                    }
                    catch (Exception ex)
                    {
                        this.excecutionErrorSummery(temp.deviceId, "reactine energy");
                    }
                        

                }
                catch { continue; }
            }
        }
        public static Dictionary<string, string> InterpretValues=new Dictionary<string,string>();
       
        private void filewriter(currentlistTowriteFile temp,string ct)
        {
            try
            {
                String filePath;
                if (Form1.CallibrationDataFolderPath == "None")
                {
                    string workingDir = Directory.GetCurrentDirectory().Replace("WindowsFormsApplication5\\bin\\Release", "").Replace("WindowsFormsApplication5\\bin\\Debug", "");
                    filePath = workingDir + "OffsetValues\\" + temp.deviceID + ".csv";
                }
                else
                {
                    filePath = Form1.CallibrationDataFolderPath + "\\" + temp.deviceID + ".csv";
                }
                String tempOffsetfilePath="F:\\calibrationFiles\\temporyOffsetFiles\\" + temp.deviceID + ".csv";
                if (!File.Exists(filePath))
                {
                    File.Create(filePath).Close();
                }

                int valueA = 1000000;
                int valueB = 1000000;
                int valueC = 1000000;
                try
                {
                    valueA = decimal.ToInt32(temp.IRMSOSPhaseA);
                    if (valueA > 2047)
                        valueA = 2047;
                    else if (valueA < -2048)
                        valueA = -2048;
                    valueB = decimal.ToInt32(temp.IRMSOSPhaseB);
                    if (valueB > 2047)
                        valueB = 2047;
                    else if (valueB < -2048)
                        valueB = -2048;
                    valueC = decimal.ToInt32(temp.IRMSOSPhaseC);
                    if (valueC > 2047)
                        valueC = 2047;
                    else if (valueC < -2048)
                        valueC = -2048;
                }
                catch (Exception ex) { }

                Equations.InterpretValues.Add(temp.deviceID, ct + "_," + valueA + "," + valueB + "," + valueC + "");
                //if (!File.Exists(tempOffsetfilePath))
                //{
                //    File.Create(tempOffsetfilePath).Close();
                //}

                String delimiter = ",";
                String[][] output = new String[][]{
            new String[]{temp.deviceID,temp.port,"phaseA","PhaseB","PhaseC"+"\n\n","Current calibration\n\n","fullscaleCurrent/500",temp.fulscaleCurrentdiv500.ToString(),temp.fulscaleCurrentdiv500.ToString(),temp.fulscaleCurrentdiv500.ToString()+"\n","Corresponding IRMS reg value",temp.PhaseARegValuForFcurrentDiv50.ToString(),temp.PhaseBRegValuForFcurrentDiv50.ToString(),temp.PhaseCRegValuForFcurrentDiv50.ToString()+"\n","test current",temp.testCurrent.ToString(),temp.testCurrent.ToString(),temp.testCurrent.ToString()+"\n"+" ","Corresponding IRMS reg Value",temp.RegvaluePhaseAforTestCurrent.ToString(),temp.RegvaluePhaseBforTestCurrent.ToString(),temp.RegvaluePhaseBforTestCurrent.ToString()+"\n","IRMSOS",(temp.IRMSOSPhaseA).ToString(),(temp.IRMSOSPhaseB).ToString(),(temp.IRMSOSPhaseC).ToString()+"\n"} /*add the values that you want inside a csv file. Mostly this function can be used in a foreach loop.*/
            };
                int length = output.GetLength(0);
                StringBuilder sb = new StringBuilder();
                for (int index = 0; index < length; index++)
                    sb.AppendLine(String.Join(delimiter, output[index]));
                File.AppendAllText(filePath, sb.ToString());

                //write to the tempory file
                //String delimiter2 = ",";
                //String[][] output2 = new String[][]{
                //new String[]{temp.deviceID+"\n","IRMSOS",(temp.IRMSOSPhaseA).ToString(),(temp.IRMSOSPhaseB).ToString(),(temp.IRMSOSPhaseC).ToString()+"\n"} /*add the values that you want inside a csv file. Mostly this function can be used in a foreach loop.*/
                //};
                //int length2 = output2.GetLength(0);
                //StringBuilder sb2 = new StringBuilder();
                //for (int index = 0; index < length; index++)
                //    sb2.AppendLine(String.Join(delimiter2, output2[index]));
                //File.AppendAllText(tempOffsetfilePath, sb2.ToString());
                




            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        public void filewriterforVoltage(VoltagelistTowriteFile temp)
        {
            try
            {
                String filePath;
                if (Form1.CallibrationDataFolderPath == "None")
                {
                    string workingDir = Directory.GetCurrentDirectory().Replace("WindowsFormsApplication5\\bin\\Release", "").Replace("WindowsFormsApplication5\\bin\\Debug", "");
                    filePath = workingDir+"OffsetValues\\" + temp.deviceID + ".csv";
                }
                else filePath = Form1.CallibrationDataFolderPath + "\\" + temp.deviceID + ".csv";
                String tempOffsetfilePath = "F:\\calibrationFiles\\temporyOffsetFiles\\" + temp.deviceID + ".csv";
               
                if (!File.Exists(filePath))
                {
                    File.Create(filePath).Close();
                }
                //if (!File.Exists(tempOffsetfilePath))
                //{
                //    File.Create(tempOffsetfilePath).Close();
                //}
                int valueA=1000000;
                int valueB = 1000000;
                int valueC = 1000000;
                try
                {
                    valueA = decimal.ToInt32(temp.PhaseAVRMSOS);
                    if (valueA > 2047)
                        valueA = 2047;
                    else if (valueA < -2048)
                        valueA = -2048;
                     valueB = decimal.ToInt32(temp.PhaseBVRMSOS);
                    if (valueB > 2047)
                        valueB = 2047;
                    else if (valueB < -2048)
                        valueB = -2048;
                     valueC = decimal.ToInt32(temp.PhaseCVRMSOS);
                    if (valueC > 2047)
                        valueC = 2047;
                    else if (valueC < -2048)
                        valueC = -2048;
                }
                catch (Exception ex) { }
                try
                {
                    string value = Equations.InterpretValues[temp.deviceID];
                    Equations.InterpretValues.Remove(temp.deviceID);
                    Equations.InterpretValues.Add(temp.deviceID, value + "_," + valueA + "," + valueB + "," + valueC + "");
                }
                catch (Exception ec) { }

                String delimiter = ",";
                String[][] output = new String[][]{
            new String[]{" ","Voltage Calibration"+"\n\n","Fullscale Voltage / 20",temp.fulscaleVoltageDiv20.ToString(),temp.fulscaleVoltageDiv20.ToString(),temp.fulscaleVoltageDiv20.ToString()+"\n","Corresponding VRMS reg Value",temp.PhaseARegValuForFVoltageDiv20.ToString(),temp.PhaseBRegValuForFVoltageDiv20.ToString(),temp.PhaseCRegValuForFVoltageDiv20.ToString()+"\n","test Voltage",temp.testVoltage.ToString(),temp.testVoltage.ToString(),temp.testVoltage.ToString()+"\n","Corresponding VRMS reg Value",temp.testRegvalForVoltagePhaseA.ToString(),temp.testRegvalForVoltagePhaseB.ToString(),temp.testRegvalForVoltagePhaseC.ToString()+"\n","VRMSOS",(temp.PhaseAVRMSOS).ToString(),(temp.PhaseBVRMSOS).ToString(),(temp.PhaseCVRMSOS).ToString()+"\n"} /*add the values that you want inside a csv file. Mostly this function can be used in a foreach loop.*/
            };
                int length = output.GetLength(0);
                StringBuilder sb = new StringBuilder();
                for (int index = 0; index < length; index++)
                    sb.AppendLine(String.Join(delimiter, output[index]));
                File.AppendAllText(filePath, sb.ToString());

                //String[][] output2 = new String[][]{
                //new String[]{temp.deviceID+"\n","VRMSOS",(temp.PhaseAVRMSOS).ToString(),(temp.PhaseBVRMSOS).ToString(),(temp.PhaseCVRMSOS).ToString()+"\n"} /*add the values that you want inside a csv file. Mostly this function can be used in a foreach loop.*/
                //};
                //int length2 = output2.GetLength(0);
                //StringBuilder sb2 = new StringBuilder();
                //for (int index = 0; index < length; index++)
                //    sb2.AppendLine(String.Join(delimiter, output2[index]));
                //File.AppendAllText(tempOffsetfilePath, sb2.ToString());


            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            // File.


        }

        public void filewriterforPhase(PhaselistTowriteFile temp)
        {
            try
            {
                String filePath;
                if (Form1.CallibrationDataFolderPath == "None")
                {
                    string workingDir = Directory.GetCurrentDirectory().Replace("WindowsFormsApplication5\\bin\\Release", "").Replace("WindowsFormsApplication5\\bin\\Debug", "");
                    filePath =workingDir+ "OffsetValues\\" + temp.deviceID + ".csv";
                }
                else filePath = Form1.CallibrationDataFolderPath + "\\" + temp.deviceID + ".csv";
                if (!File.Exists(filePath))
                {
                    File.Create(filePath).Close();
                }

                int valueA = 1000000;
                int valueB = 1000000;
                int valueC = 1000000;
                try
                {
                    valueA = decimal.ToInt32(temp.xPHCALA);
                    if (valueA > 64)
                        valueA = 64;
                    else if (valueA < -65)
                        valueA = -65;
                    valueB = decimal.ToInt32(temp.xPHCALB);
                    if (valueB > 64)
                        valueB = 64;
                    else if (valueB < -65)
                        valueB = -65;
                    valueC = decimal.ToInt32(temp.xPHCALC);
                    if (valueC > 64)
                        valueC = 64;
                    else if (valueC < -65)
                        valueC = -65;
                }
                catch (Exception ex) { }
                try
                {
                    string value = Equations.InterpretValues[temp.deviceID];
                    Equations.InterpretValues.Remove(temp.deviceID);
                    Equations.InterpretValues.Add(temp.deviceID, value + "_," + valueA + "," + valueB + "," + valueC + ",");
                }
                catch (Exception ex) { }

                String delimiter = ",";
                String[][] output = new String[][]{
            new String[]{" ","Phase Calibration"+"\n\n","Power at pf 1 at Itest and Vnom",temp.Power_at_pf_1_at_Itest_and_Vnom_phaseA.ToString(),temp.Power_at_pf_1_at_Itest_and_Vnom_phaseB.ToString(),temp.Power_at_pf_1_at_Itest_and_Vnom_phaseC.ToString()+"\n","Power at pf x at Itest and Vnom",temp.Power_at_pf_x_at_Itest_and_Vnom_phaseA.ToString(),temp.Power_at_pf_x_at_Itest_and_Vnom_phaseB.ToString(),temp.Power_at_pf_x_at_Itest_and_Vnom_phaseC.ToString()+"\n","Used Power factor",temp.Used_Power_factor.ToString(),temp.Used_Power_factor.ToString(),temp.Used_Power_factor.ToString()+"\n","Error",temp.ErrorA.ToString(),temp.ErrorB.ToString(),temp.ErrorC.ToString()+"\n","Phase Error",temp.Phase_ErrorA.ToString(),temp.Phase_ErrorB.ToString(),temp.Phase_ErrorC.ToString()+"\n","xPHCAL",(temp.xPHCALA).ToString(),(temp.xPHCALB).ToString(),(temp.xPHCALC).ToString()} /*add the values that you want inside a csv file. Mostly this function can be used in a foreach loop.*/
            };
                int length = output.GetLength(0);
                StringBuilder sb = new StringBuilder();
                for (int index = 0; index < length; index++)
                    sb.AppendLine(String.Join(delimiter, output[index]));
                File.AppendAllText(filePath, sb.ToString());
                // File.
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }


        }

        public void fileWriterForActiveEnergy(energylistWriteFile temp) {
            try
            {
                String filePath;
                if (Form1.CallibrationDataFolderPath == "None")
                {
                    string workingDir = Directory.GetCurrentDirectory().Replace("WindowsFormsApplication5\\bin\\Release", "").Replace("WindowsFormsApplication5\\bin\\Debug", "");
                    filePath = workingDir+"OffsetValues\\" + temp.deviceId + ".csv";
                }
                else filePath = Form1.CallibrationDataFolderPath + "\\" + temp.deviceId + ".csv";
                if (!File.Exists(filePath))
                {
                    File.Create(filePath).Close();
                }

                int valueA = 1000000;
                int valueB = 1000000;
                int valueC = 1000000;
                try
                {
                    valueA = decimal.ToInt32(decimal.Parse(temp.xWATTOSA));
                    if (valueA > 2047)
                        valueA = 2047;
                    else if (valueA < -2048)
                        valueA = -2048;
                    valueB = decimal.ToInt32(decimal.Parse(temp.xWATTOSB));
                    if (valueB > 2047)
                        valueB = 2047;
                    else if (valueB < -2048)
                        valueB = -2048;
                    valueC = decimal.ToInt32(decimal.Parse(temp.xWATTOSC));
                    if (valueC > 2047)
                        valueC = 2047;
                    else if (valueC < -2048)
                        valueC = -2048;
                }
                catch (Exception ex) { }
                try
                {
                    string value = Equations.InterpretValues[temp.deviceId];
                    Equations.InterpretValues.Remove(temp.deviceId);
                    Equations.InterpretValues.Add(temp.deviceId, value + "_," + valueA + "," + valueB + "," + valueC + "");
                }
                catch (Exception ex) { }

                String delimiter = ",";
                String[][] output = new String[][]{
            new String[]{"\n","Active Energy"+"\n\n","Wt1600MeasureFirstEnergy",temp.energfrimWt16001st.ToString(),temp.energfrimWt16001st.ToString(),temp.energfrimWt16001st.ToString()+"\n","Wt1600MeasureSecondEnergy",temp.energfrimWt16002nd.ToString(),temp.energfrimWt16002nd.ToString(),temp.energfrimWt16002nd.ToString()+"\n","Corresponding reg value for 1st",temp.PhaseACorrespondingWATTHRregValueForFirst.ToString(),temp.PhaseBCorrespondingWATTHRregValueForFirst.ToString(),temp.PhaseCCorrespondingWATTHRregValueForFirst.ToString()+"\n","Line Cycles",temp.LineCycle1.ToString(),temp.LineCycle1.ToString(),temp.LineCycle1.ToString()+"\n","Reg values For second intergration",temp.TestPhaseACorrespondingWATTHRregValue.ToString(),temp.TestPhaseBCorrespondingWATTHRregValue.ToString(),temp.TestPhaseCCorrespondingWATTHRregValue.ToString()+"\n","Line cycle for 2nd",temp.TestLineCycle.ToString(),temp.TestLineCycle.ToString(),temp.TestLineCycle.ToString()+"\n","Offset",temp.OffsetA.ToString(),temp.OffsetB.ToString(),temp.OffsetC.ToString()+"\n","xWATTOS",temp.xWATTOSA.ToString(),temp.xWATTOSB.ToString(),temp.xWATTOSC.ToString()} /*add the values that you want inside a csv file. Mostly this function can be used in a foreach loop.*/
            };
                int length = output.GetLength(0);
                StringBuilder sb = new StringBuilder();
                for (int index = 0; index < length; index++)
                    sb.AppendLine(String.Join(delimiter, output[index]));
                File.AppendAllText(filePath, sb.ToString());
                // File.
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void fileWriterForAReactiveEnergy(ReactiveEnergylistWriteFile temp)
        {
            try
            {
                String filePath;
                if (Form1.CallibrationDataFolderPath == "None")
                {
                    string workingDir = Directory.GetCurrentDirectory().Replace("WindowsFormsApplication5\\bin\\Release", "").Replace("WindowsFormsApplication5\\bin\\Debug", "");
                    filePath = workingDir+"OffsetValues\\" + temp.deviceId + ".csv";
                }
                else filePath = Form1.CallibrationDataFolderPath + "\\" + temp.deviceId + ".csv";
                if (!File.Exists(filePath))
                {
                    File.Create(filePath).Close();
                }
                int valueA = 1000000;
                int valueB = 1000000;
                int valueC = 1000000;
                try
                {
                    valueA = decimal.ToInt32(decimal.Parse(temp.xWATTOSA));
                    if (valueA > 2047)
                        valueA = 2047;
                    else if (valueA < -2048)
                        valueA = -2048;
                    valueB = decimal.ToInt32(decimal.Parse(temp.xWATTOSB));
                    if (valueB > 2047)
                        valueB = 2047;
                    else if (valueB < -2048)
                        valueB = -2048;
                    valueC = decimal.ToInt32(decimal.Parse(temp.xWATTOSC));
                    if (valueC > 2047)
                        valueC = 2047;
                    else if (valueC < -2048)
                        valueC = -2048;
                }
                catch (Exception ex) { }
                try
                {
                    string value = Equations.InterpretValues[temp.deviceId] + "_," + valueA + "," + valueB + "," + valueC + "";
                    Equations.InterpretValues.Remove(temp.deviceId);
                    this.writeFinalOutputs(value, temp.deviceId);
                }
                catch (Exception ex) { }
                String delimiter = ",";
                String[][] output = new String[][]{
            new String[]{"\n","Reactive Energy"+"\n\n","Wt1600MeasureFirstEnergy",temp.energfrimWt16001st.ToString(),temp.energfrimWt16001st.ToString(),temp.energfrimWt16001st.ToString()+"\n","Wt1600MeasureSecondEnergy",temp.energfrimWt16002nd.ToString(),temp.energfrimWt16002nd.ToString(),temp.energfrimWt16002nd.ToString()+"\n","Corresponding reg value for 1st",temp.PhaseACorrespondingWATTHRregValueForFirst.ToString(),temp.PhaseBCorrespondingWATTHRregValueForFirst.ToString(),temp.PhaseCCorrespondingWATTHRregValueForFirst.ToString()+"\n","Line Cycles",temp.LineCycle1.ToString(),temp.LineCycle1.ToString(),temp.LineCycle1.ToString()+"\n","Reg values For second intergration",temp.TestPhaseACorrespondingWATTHRregValue.ToString(),temp.TestPhaseBCorrespondingWATTHRregValue.ToString(),temp.TestPhaseCCorrespondingWATTHRregValue.ToString()+"\n","Line cycle for 2nd",temp.TestLineCycle.ToString(),temp.TestLineCycle.ToString(),temp.TestLineCycle.ToString()+"\n","Offset",temp.OffsetA.ToString(),temp.OffsetB.ToString(),temp.OffsetC.ToString()+"\n","xWATTOS",temp.xWATTOSA.ToString(),temp.xWATTOSB.ToString(),temp.xWATTOSC.ToString()} /*add the values that you want inside a csv file. Mostly this function can be used in a foreach loop.*/
            };
                int length = output.GetLength(0);
                StringBuilder sb = new StringBuilder();
                for (int index = 0; index < length; index++)
                    sb.AppendLine(String.Join(delimiter, output[index]));
                File.AppendAllText(filePath, sb.ToString());
                // File.
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }
        
        private void writeFinalOutputs(string InterpretValues,string id)
        {
            string[] chunks=InterpretValues.Split('_');
            try
            {
                String filePath;
                if (Form1.CallibrationDataFolderPath == "None")
                {
                    string workingDir = Directory.GetCurrentDirectory().Replace("WindowsFormsApplication5\\bin\\Release", "").Replace("WindowsFormsApplication5\\bin\\Debug", "");
                    filePath = workingDir + "OffsetValues\\P" + id +"-"+chunks[0]+ "A.txt";
                }
                else filePath = Form1.CallibrationDataFolderPath + "\\P" + id + "-" + chunks[0] + "A.txt";
                if (!File.Exists(filePath))
                {
                    File.Create(filePath).Close();
                }
                using (StreamWriter sw = File.AppendText(filePath))
                {
                    try
                    {
                        sw.WriteLine(chunks[2] + chunks[1] + chunks[4] + chunks[5] + chunks[3]);
                    }
                    catch (IndexOutOfRangeException ex)
                    {
                        sw.WriteLine("Error in data!");
                    }
                }
            }

            catch (Exception ex)
            {
            }


        }

        public class currentlistTowriteFile
        {
            public decimal fulscaleCurrentdiv500;
            public decimal PhaseARegValuForFcurrentDiv50;
            public decimal PhaseBRegValuForFcurrentDiv50;
            public decimal PhaseCRegValuForFcurrentDiv50;
            public decimal testCurrent;
            public decimal RegvaluePhaseAforTestCurrent;
            public decimal RegvaluePhaseBforTestCurrent;
            public decimal RegvaluePhaseCforTestCurrent;
            public decimal IRMSOSPhaseA;
            public decimal IRMSOSPhaseB;
            public decimal IRMSOSPhaseC;
            public String  deviceID;
            public String port;
               
        }

        public class VoltagelistTowriteFile
        {
            public String deviceID;

            public decimal fulscaleVoltageDiv20;
            public decimal PhaseARegValuForFVoltageDiv20;
            public decimal PhaseBRegValuForFVoltageDiv20;
            public decimal PhaseCRegValuForFVoltageDiv20;

            public decimal testVoltage;
            public decimal testRegvalForVoltagePhaseA;
            public decimal testRegvalForVoltagePhaseB;
            public decimal testRegvalForVoltagePhaseC;

            public decimal PhaseAVRMSOS;
            public decimal PhaseBVRMSOS;
            public decimal PhaseCVRMSOS;

        }

        public class PhaselistTowriteFile 
        {
        public decimal Power_at_pf_1_at_Itest_and_Vnom_phaseA;
        public decimal Power_at_pf_1_at_Itest_and_Vnom_phaseB;
        public decimal Power_at_pf_1_at_Itest_and_Vnom_phaseC;

        public decimal Power_at_pf_x_at_Itest_and_Vnom_phaseA;
        public decimal Power_at_pf_x_at_Itest_and_Vnom_phaseB;
        public decimal Power_at_pf_x_at_Itest_and_Vnom_phaseC;
        public double Used_Power_factor;

        public double ErrorA, ErrorB, ErrorC;
        public decimal Phase_ErrorA, Phase_ErrorB, Phase_ErrorC;
        public decimal xPHCALA, xPHCALB, xPHCALC;
        public String deviceID;

        }

        public class energylistWriteFile
        {
            public String deviceId;
            public String PhaseACorrespondingWATTHRregValueForFirst;
            public String PhaseBCorrespondingWATTHRregValueForFirst;
            public String PhaseCCorrespondingWATTHRregValueForFirst;
            public String energfrimWt16001st;
            public String LineCycle1;
            public String OffsetA;
            public String OffsetB;
            public String OffsetC;
            public String xWATTOSA;
            public String xWATTOSB;
            public String xWATTOSC;
          
            public String TestPhaseACorrespondingWATTHRregValue;
            public String TestPhaseBCorrespondingWATTHRregValue;
            public String TestPhaseCCorrespondingWATTHRregValue;
            public String energfrimWt16002nd;
            public String TestLineCycle;
              
        }

        public class ReactiveEnergylistWriteFile
        {
            public String deviceId;
            public String PhaseACorrespondingWATTHRregValueForFirst;
            public String PhaseBCorrespondingWATTHRregValueForFirst;
            public String PhaseCCorrespondingWATTHRregValueForFirst;
            public String energfrimWt16001st;
            public String LineCycle1;

            public String TestPhaseACorrespondingWATTHRregValue;
            public String TestPhaseBCorrespondingWATTHRregValue;
            public String TestPhaseCCorrespondingWATTHRregValue;
            public String energfrimWt16002nd;
            public String TestLineCycle;

            public String OffsetA;
            public String OffsetB;
            public String OffsetC;
            public String xWATTOSA;
            public String xWATTOSB;
            public String xWATTOSC;

        }




        public void excecutionErrorSummery(string id, string field)
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
                sw.WriteLine("Parameters for :" + field + " coudn't write back to the device (:" + id + ")!");
            }
        }


       
    }
}
