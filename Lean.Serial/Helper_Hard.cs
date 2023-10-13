using System;
using System.Collections.Generic;
using System.Management;
using System.Net;

namespace Lean.Serial
{
    class Helper_Hard
    {
        //1.获取CPU序列号代码  
        public static string GetCpuID()
        {
            try
            {
                string cpuInfo = "";//cpu序列号  
                ManagementClass mc = new ManagementClass("Win32_Processor");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    cpuInfo = mo.Properties["ProcessorId"].Value.ToString();
                }
                moc = null;
                mc = null;
                return cpuInfo;
            }
            catch
            {
                return "unknow";
            }
        }
        //2.获取网卡硬件地址  
        public static string GetMacAddress()
        {
            try
            {
                string mac = "";
                ManagementClass mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    if ((bool)mo["IPEnabled"] == true)
                    {
                        mac = mo["MacAddress"].ToString();
                        break;
                    }
                }
                moc = null;
                mc = null;
                return mac;
            }
            catch
            {
                return "unknow";
            }
        }
        //3.获取硬盘ID  
        public static string GetDiskID()
        {
            try
            {
                String HDid = "";
                ManagementClass mc = new ManagementClass("Win32_DiskDrive");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    HDid = (string)mo.Properties["Model"].Value;
                }
                moc = null;
                mc = null;
                return HDid;
            }
            catch
            {
                return "unknow";
            }
        }
        //4.获取IP地址  
        public static string GetIPAddress()
        {
            try
            {
                string st = "";
                ManagementClass mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    if ((bool)mo["IPEnabled"] == true)
                    {
                        //st=mo["IpAddress"].ToString();  
                        System.Array ar;
                        ar = (System.Array)(mo.Properties["IpAddress"].Value);
                        st = ar.GetValue(0).ToString();
                        break;
                    }
                }
                moc = null;
                mc = null;
                return st;
            }
            catch
            {
                return "unknow";
            }
        }
        // 5.操作系统的登录用户名  
        public static string GetUserName()
        {
            //string st;
            try
            {
                //string un = "";
                //st = Environment.UserName;
                return System.Environment.UserName;
                //return un;
            }
            catch
            {
                return "unknow";
            }
        }
        //6.获取计算机名 
        public static string GetComputerName()
        {
            try
            {
                return System.Environment.MachineName;
            }
            catch
            {
                return "unknow";
            }
        }
        //7. PC类型  
        public static string GetSystemType()
        {
            try
            {
                string st = "";
                ManagementClass mc = new ManagementClass("Win32_ComputerSystem");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    st = mo["SystemType"].ToString();
                }
                moc = null;
                mc = null;
                return st;
            }
            catch
            {
                return "unknow";
            }
        }
        // 8.物理内存  
        public static string GetTotalPhysicalMemory()
        {

            try
            {
                string st = "";

                ManagementClass mc = new ManagementClass("Win32_ComputerSystem");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    st = mo["TotalPhysicalMemory"].ToString();

                }
                moc = null;
                mc = null;
                return st;
            }

            catch
            {
                return "unknow";
            }

        }
    }
}
