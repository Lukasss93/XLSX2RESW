using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using XLSX2RESW.Classes;

namespace XLSX2RESW
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            //no double instance!
            if(Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1)
            {
                Current.Shutdown();
                return;
            }

            //check parameters
            if(e.Args.Length > 0)
            {                
                var files = e.Args;

                //convert files
                Elaborator.Convert(files);

                //exit from application
                Current.Shutdown();
                return;
            }

        }
    }
}
