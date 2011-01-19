using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;
using System.IO;
using System.Diagnostics;

namespace CommitActions
{
    class Program
    {
        internal static List<string> VSExeFolders;

        static void Main(string[] args)
        {
            //load visual studio exe locations           
            RegistryKey visualStudioRootReg = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\VisualStudio");
            VSExeFolders = visualStudioRootReg.GetSubKeyNames().ToList().Where(i => visualStudioRootReg.OpenSubKey(i).GetValue("InstallDir") != null).Select(i => Convert.ToString(visualStudioRootReg.OpenSubKey(i).GetValue("InstallDir"))).ToList();
            VSExeFolders = VSExeFolders.Where(i => File.Exists(i.Trim('\\') + "\\" + "devenv.exe")).Select(i => i.Trim('\\') + "\\" + "devenv.exe").ToList();
            VSExeFolders.Distinct().ToList().ForEach(
                i =>
                {
                    Process process = new Process();
                    process.StartInfo = new ProcessStartInfo(i, "/resetaddin ILViewer.Connect");
                    process.Start();
                }
                );
        }

    }
}
