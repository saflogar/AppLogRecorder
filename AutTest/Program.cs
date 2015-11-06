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
        [DllImport("user32", EntryPoint = "FindWindow", CharSet = CharSet.Auto)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    

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
            
            }catch(Exception ex)
            {
                Console.WriteLine("Fatal Error" +  ex.Message);
            }
           
        }
        }
    }

