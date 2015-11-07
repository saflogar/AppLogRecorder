using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Data;


namespace AutTest
{
    class Program
    {
        private String _MainMenuWindName = "";

        [DllImport("user32", EntryPoint = "FindWindow", CharSet = CharSet.Auto)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32", EntryPoint = "FindWindowEx", CharSet = CharSet.Auto)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent,IntPtr hwndChildAfter,string lpzClass, string lpzWindow);

        [DllImport("user32", EntryPoint = "GetWindowText", CharSet = CharSet.Auto)]
        static extern bool GetWindowText(IntPtr hWnd, string lpString, int nMaxcount);

        [DllImport("user32", EntryPoint = "PostMessage", CharSet = CharSet.Auto)]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

        uint WM_LBUTTONDOWN = 0x0201;
        uint WM_LBUTTONUP = 0x202;

        private void ExecuteCommand(String command, string[] args, IntPtr mwh) 
        {
            switch (command)
            {   
                case "NAVTO":
                    if (args[0].Equals("FORM")) 
                    {
                        string wText = null; 
                        GetWindowText(mwh,wText,_MainMenuWindName.Length);
                        if (wText.Equals(_MainMenuWindName))
                        {
                            IntPtr frmCtrWH = FindWindowByIndex(mwh,int.Parse(args[1]));
                        }
                    }
                    else if (args[0].Equals("TAB"))
                    {
                    }
                break;
                case "Click":

                default:
                    break;
            }
        }

        private IntPtr FindWindowByIndex(IntPtr hwndParent, int index) 
        {
            if (index == 0)
            {
                return hwndParent;
            }
            else 
            {
                int ct = 0;
                IntPtr result = IntPtr.Zero;
                do
                {
                    result = FindWindowEx(hwndParent, result, null, null);
                } while (ct < index && result != IntPtr.Zero);
                return result;
            }

        }

        static void Main(string[] args)
        {
            try 
            {
                String path = "C:\\Users\\VBW7\\Documents\\Visual Studio 2013\\Projects\\WindowsFormsApplication1\\WindowsFormsApplication1\\bin\\Debug\\WindowsFormsApplication1.exe";
                String connString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
"Data Source=C:\\Results\\results.xls;" +
"Extended Properties=\"Excel 8.0;HDR=YES;IMEX=0\"";

                Process p = Process.Start(path);

                Reader r = new Reader(connString);

                DataTable dt = r.Dt;

                foreach (DataRow dr in dt.Rows)
                {

                }

                IntPtr mwh = IntPtr.Zero;
                bool formFound = false;
                int attempts = 0;

                while (!formFound && attempts < 50)
                {
                    if (mwh == IntPtr.Zero)
                    {
                        Console.WriteLine("Form not yet found");
                        Thread.Sleep(100);
                        ++attempts;
                        mwh = FindWindow(null, "Form1");
                    }
                    else
                    {
                        Console.WriteLine("Form has been found");
                        formFound = true;
                    }
                }

                if (mwh == IntPtr.Zero)
                    throw new Exception("Could not find main window");
                else 
                {
                    switch ()
                    {
                    }
                
                }
            
            }catch(Exception ex)
            {
                Console.WriteLine("Fatal Error" +  ex.Message);
            }
           
        }
        }
    }

