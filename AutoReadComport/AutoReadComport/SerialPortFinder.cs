using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace AutoReadComport
{
    public class SerialPortFinder
    {
         public string[] FindAllPorts()
        {
            List<string> ports = new List<string>();
            foreach (ManagementObject item in FindPorts())
            {
                try
                {
                    if (item["Caption"].ToString().Contains("COM"))
                    {
                        string comName = ParseNameCOMPort(item);
                        if (comName != null)
                            ports.Add(comName);
                    }
                }
                catch (Exception ex)
                {
                   
                }
            }
            return ports.ToArray();
        }

        private string ParseNameCOMPort(ManagementObject item)
        {
            string name = item["Name"].ToString();
            int startIndex = name.LastIndexOf("(");
            int endIndex = name.LastIndexOf(")");
            if (startIndex != -1 && endIndex != -1)
            {
                name = name.Substring(startIndex + 1, endIndex - startIndex - 1);
                return name;
            }
            return null;

        }

        public string FindPortByDescription(string description)
        {
            foreach (ManagementObject item in FindPorts())
            {
                try
                {
                    if (item["Description"].ToString().ToLower().Equals(description.ToLower()))
                    {
                        string comName = ParseNameCOMPort(item);
                        if (comName != null)
                            return comName;
                    }
                }
                catch (Exception ex)
                {
                   
                }
            }
            return null;
        }
        public string FindPortByManufacture(string manufacturer)
        {
            foreach (ManagementObject item in FindPorts())
            {
                try
                {
                    if (item["Manufacturer"].ToString().ToLower().Equals(manufacturer.ToLower()))
                    {
                        string comName = ParseNameCOMPort(item);
                        if (comName != null)
                            return comName;
                    }
                }
                catch (Exception ex)
                {
                    
                }
            }
            return null;
        }

        public ManagementObject[] FindPorts()
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PnPEntity");
                List<ManagementObject> objects = new List<ManagementObject>();
                foreach (ManagementObject item in searcher.Get())
                {
                    objects.Add(item);
                }
                return objects.ToArray();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return new ManagementObject[] { };
            }
        }
    }
}
