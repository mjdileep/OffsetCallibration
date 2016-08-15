using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NationalInstruments.NI4882;
using System.Configuration;
using System.Threading;
using System.Diagnostics;

namespace WindowsFormsApplication5
{
   

    public class GetdataWT1600Energy 
    {

        // Program pro=new Program();    

        private Device device;

        int BoardId = 0;
        byte primaryAddress = 1;
        byte currentSecondaryAddress = 0;


        public void OpenDevice()
        {

            device = new Device(0, primaryAddress, currentSecondaryAddress);
            //Board board = new Board(i); // i is whatever board number you want

            //board.SendInterfaceClear();
            //   SetupControlState(true);

        }

        public void CloseDevice()
        {

            device.Dispose();

        }


        public string returnEnergy(string LineCircles)
        {

            string endCommand = "INTEGRATE:STOP";
           
            string setnumOfVariablesvalues = "NUM:NORM:NUM 255\n";
            string setPattern = "NUMERIC:NORMAL:PRESET 4\n";
            string AskVariablesvalues = "NUM:NORM:VAL?\n";

            string readVar = "NUM:NORM:NUM?\n";
            string resetdevice = "*rst\n";


          //  device.Write(ReplaceCommonEscapeSequences(endCommand));

            device.Write(ReplaceCommonEscapeSequences(setnumOfVariablesvalues));
            device.Write(ReplaceCommonEscapeSequences(setPattern));
            device.Write(ReplaceCommonEscapeSequences(AskVariablesvalues));



            string readbuffer = InsertCommonEscapeSequences(device.ReadString(2350));
            string[] words = readbuffer.Split(',');
            string ENERGY = words[14];
            

            //  int j = 0;
            // while (j < 255) { Console.WriteLine(j+1+"="+words[j]); j++; }

            Console.WriteLine(words.Count());
            //int i = 0;
            //while (words[i] != "\n") { Console.WriteLine(i+1 + "=" + words[i] + "\n"); i++; }
            device.Write(ReplaceCommonEscapeSequences(resetdevice));
            return ENERGY;
        }

        public void startIntegration(string time,string sensitive) 
        {   //IF NEED CHECK THE state and stop intgration if not had stopeed last time // title 
            string status = "INTEGRATE:STATE?";
            string reset = "INTEGRATE:RESET";
            string startInt = "INTEGRATE:START";
            string temp = "INTEGRATE:TIMER " + time;
            string setTimer = temp;

            if (sensitive == "high") { device.Write(ReplaceCommonEscapeSequences("INPUT:CURRENT:RANGE:ALL 5A")); }
            device.Write(ReplaceCommonEscapeSequences(reset));
            device.Write(ReplaceCommonEscapeSequences(setTimer));
            device.Write(ReplaceCommonEscapeSequences(startInt));

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




    }





}
