using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReadComport
{
    class Program
    {
        static void Main(string[] args)
        {
            var serialPort = new SerialPortFinder();

            //foreach (var item in serialPort.FindAllPorts())
            //{
            //    Console.WriteLine(item);
            //}
            // string comDevice = serialPort.FindPortByDescription("MOXA USB Serial Port");
            string comDevice = serialPort.FindPortByManufacture("Moxa Inc.");
            if (comDevice != null)
                Console.WriteLine(comDevice);
            Console.Read();
        }
    }
}
