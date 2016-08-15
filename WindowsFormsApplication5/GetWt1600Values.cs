using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NationalInstruments.NI4882;
using System.IO;

namespace WindowsFormsApplication5
{
    public class GetdataWT1600
    {

        // Program pro=new Program();    

        private Device device;

        int BoardId = 0;
        byte primaryAddress = 1;
        byte currentSecondaryAddress = 0;


        public void OpenDevice()
        {

            device = new Device(0, primaryAddress, currentSecondaryAddress);
            //   SetupControlState(true);

        }

        public void CloseDevice()
        {

            device.Dispose();

        }


        public string writeAndtakeAlltheParameters(string type)
        {
            if (type == "firstCurrent") 
            {
                device.Write(ReplaceCommonEscapeSequences("INPUT:CURRENT:RANGE:ALL 5A"));
            }


            String command = "NUM:NORM:NUM 90\n";
           // device.Write(ReplaceCommonEscapeSequences(ibpct));
            device.Write(ReplaceCommonEscapeSequences(command));
            // device.Write(ReplaceCommonEscapeSequences("*CLS/n"));
            device.Write(ReplaceCommonEscapeSequences("NUM:NORM:VAL?\n"));
            String readbuffer;
            readbuffer = InsertCommonEscapeSequences(device.ReadString(1024));
            // handleCurrent(readbuffer);
            List<string> substrings = readbuffer.Split(',').ToList();
            //  Console.WriteLine(substrings.Count);
            List<decimal> VRMS = new List<decimal>();
            List<decimal> IRMS = new List<decimal>();
            List<decimal> PowerFactor = new List<decimal>();
            //decimal temp_value = 0;
            for (int i = 0; i < substrings.Count; i = i + 15)
            {
                try
                {
                    VRMS.Add(Decimal.Parse(substrings[i], System.Globalization.NumberStyles.Float));
                }
                catch (Exception ex) { this.excecutionErrorSummery("VRMS", substrings[i]); }
            }

            for (int i = 4; i < substrings.Count; i = i + 15)
            {
                try
                {
                    IRMS.Add(Decimal.Parse(substrings[i], System.Globalization.NumberStyles.Float));
                }
                catch (Exception ex) { this.excecutionErrorSummery("IRMS", substrings[i]); }
            }
          //  if (type != "firstCurrent")
         //   {

                for (int i = 11; i < substrings.Count; i = i + 15)
                { //deviding each elements 8th and 9th elements (from 0 to 14)
                    decimal p = 0;
                     try
                    {
                        p = Decimal.Parse(substrings[i], System.Globalization.NumberStyles.Float);
                    }
                     catch (Exception ex) { this.excecutionErrorSummery("PowerFactor", substrings[i]); }
                    PowerFactor.Add(p);
                }
          //  }

            decimal averagedVrms = 0; decimal averagedIrms = 0; decimal averagedPowerFactor = 0;
            for (int i = 0; i < VRMS.Count; i++)
            {
                averagedVrms = averagedVrms + VRMS[i];
            }


            for (int i = 0; i < VRMS.Count; i++)
            {
                averagedIrms = averagedIrms + IRMS[i];
            }

          //  if (type != "firstCurrent")
        //    {

                for (int i = 0; i < PowerFactor.Count; i++)
                {
                    averagedPowerFactor = averagedPowerFactor + PowerFactor[i];
                }
        //    }

            averagedVrms = averagedVrms / 6;
            averagedIrms = averagedIrms / 6;
            averagedPowerFactor = averagedPowerFactor / 6;

            device.Write(ReplaceCommonEscapeSequences("*Rst\n"));

       //     if (type == "firstCurrent")
       //     {
       //         return averagedIrms + "," + averagedVrms;
        //    }
        //    else
        //    {
                return averagedIrms + "," + averagedVrms + "," + averagedPowerFactor;
        //    }
        }

        public string writeAndtakeAlltheParametersForVoltage()
        {
             String command = "NUM:NORM:NUM 90\n";
            // device.Write(ReplaceCommonEscapeSequences(ibpct));
            device.Write(ReplaceCommonEscapeSequences(command));
            // device.Write(ReplaceCommonEscapeSequences("*CLS/n"));
            device.Write(ReplaceCommonEscapeSequences("NUM:NORM:VAL?\n"));
            String readbuffer;
            readbuffer = InsertCommonEscapeSequences(device.ReadString(1024));
            // handleCurrent(readbuffer);
            List<string> substrings = readbuffer.Split(',').ToList();
            //  Console.WriteLine(substrings.Count);
            List<decimal> VRMS = new List<decimal>();
            List<decimal> IRMS = new List<decimal>();
         
            //decimal temp_value = 0;
            for (int i = 0; i < substrings.Count; i = i + 15)
            {
                  try
                    {
                     VRMS.Add(Decimal.Parse(substrings[i], System.Globalization.NumberStyles.Float));
                    }
                  catch (Exception ex) { this.excecutionErrorSummery("VRMS", substrings[i]); }
            }

            for (int i = 4; i < substrings.Count; i = i + 15)
            {
                try
                {
                IRMS.Add(Decimal.Parse(substrings[i], System.Globalization.NumberStyles.Float));
                }
                catch (Exception ex) { this.excecutionErrorSummery("IRMS", substrings[i]); }
            }

            


            decimal averagedVrms = 0; decimal averagedIrms = 0; decimal averagedPowerFactor = 0;
            for (int i = 0; i < VRMS.Count; i++)
            {
                averagedVrms = averagedVrms + VRMS[i];
            }


            for (int i = 0; i < VRMS.Count; i++)
            {
                averagedIrms = averagedIrms + IRMS[i];
            }


            
            averagedVrms = averagedVrms / 6;
            averagedIrms = averagedIrms / 6;
         

            device.Write(ReplaceCommonEscapeSequences("*Rst\n"));

            return averagedIrms + "," + averagedVrms; 
        }


        private string ReplaceCommonEscapeSequences(string s)
        {
            return s.Replace("\\n", "\n").Replace("\\r", "\r");
        }

        private string InsertCommonEscapeSequences(string s)
        {
            return s.Replace("\n", "").Replace("\r", "\\r");
            // return s;
        }

        public void excecutionErrorSummery(string value, string field)
        {
            String filePath;
            if (Form1.CallibrationDataFolderPath == "None")
            {
                string workingDir = Directory.GetCurrentDirectory().Replace("WindowsFormsApplication5\\bin\\Release", "").Replace("WindowsFormsApplication5\\bin\\Debug", "");
                filePath = workingDir + "OffsetValues\\excecutionErrorSummeryWT1600.txt";
            }
            else filePath = Form1.CallibrationDataFolderPath + "\\excecutionErrorSummeryWT1600.txt";
            if (!File.Exists(filePath))
            {
                File.Create(filePath).Close();
            }
            using (StreamWriter sw = File.AppendText(filePath))
            {
                sw.WriteLine("Field :" + field + " from WT1600 :" + value + " was recieved!");
            }
        }


    }

}
