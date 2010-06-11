using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Resources;
using System.Reflection;
using System.Globalization;
using Microsoft.Win32;
using System.IO;
using System.Diagnostics;

namespace ILViewer
{
    /// <summary>The object for implementing an Add-in.</summary>
    /// <seealso class='IDTExtensibility2' />
    public class Connect : IDTExtensibility2, IDTCommandTarget
    {
        /// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
        public Connect()
        {
        }

        /// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
        /// <param term='application'>Root object of the host application.</param>
        /// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
        /// <param term='addInInst'>Object representing this Add-in.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _applicationObject = (DTE2)application;
            _addInInstance = (AddIn)addInInst;
            if (connectMode == ext_ConnectMode.ext_cm_UISetup)
            {
                object[] contextGUIDS = new object[] { };
                Commands2 commands = (Commands2)_applicationObject.Commands;
                string toolsMenuName;

                try
                {
                    //If you would like to move the command to a different menu, change the word "Tools" to the 
                    //  English version of the menu. This code will take the culture, append on the name of the menu
                    //  then add the command to that menu. You can find a list of all the top-level menus in the file
                    //  CommandBar.resx.
                    string resourceName;
                    ResourceManager resourceManager = new ResourceManager("ILViewer.CommandBar", Assembly.GetExecutingAssembly());
                    CultureInfo cultureInfo = new CultureInfo(_applicationObject.LocaleID);

                    if (cultureInfo.TwoLetterISOLanguageName == "zh")
                    {
                        System.Globalization.CultureInfo parentCultureInfo = cultureInfo.Parent;
                        resourceName = String.Concat(parentCultureInfo.Name, "Tools");
                    }
                    else
                    {
                        resourceName = String.Concat(cultureInfo.TwoLetterISOLanguageName, "Tools");
                    }
                    toolsMenuName = resourceManager.GetString(resourceName);
                }
                catch
                {
                    //We tried to find a localized version of the word Tools, but one was not found.
                    //  Default to the en-US word, which may work for the current culture.
                    toolsMenuName = "Tools";
                }

                //Place the command on the tools menu.
                //Find the MenuBar command bar, which is the top-level command bar holding all the main menu items:
                Microsoft.VisualStudio.CommandBars.CommandBar menuBarCommandBar = ((Microsoft.VisualStudio.CommandBars.CommandBars)_applicationObject.CommandBars)["MenuBar"];

                //Find the Tools command bar on the MenuBar command bar:
                CommandBarControl toolsControl = menuBarCommandBar.Controls[toolsMenuName];
                CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;

                //This try/catch block can be duplicated if you wish to add multiple commands to be handled by your Add-in,
                //  just make sure you also update the QueryStatus/Exec method to include the new command names.
                try
                {
                    //Add a command to the Commands collection:
                    Command command = commands.AddNamedCommand2(_addInInstance, "ILViewer", "ILViewer", "Executes the command for ILViewer", true, 59, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                    //Add a control for the command to the tools menu:
                    if ((command != null) && (toolsPopup != null))
                    {
                        command.AddControl(toolsPopup.CommandBar, 1);
                    }
                }
                catch (System.ArgumentException)
                {
                    //If we are here, then the exception is probably because a command with that name
                    //  already exists. If so there is no need to recreate the command and we can 
                    //  safely ignore the exception.
                }
            }
        }

        /// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
        /// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
        }

        /// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />		
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom)
        {
        }

        /// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref Array custom)
        {
        }

        /// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
        /// <param term='commandName'>The name of the command to determine state for.</param>
        /// <param term='neededText'>Text that is needed for the command.</param>
        /// <param term='status'>The state of the command in the user interface.</param>
        /// <param term='commandText'>Text requested by the neededText parameter.</param>
        /// <seealso class='Exec' />
        public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                if (commandName == "ILViewer.Connect.ILViewer")
                {
                    status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                    return;
                }
            }
        }

        /// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
        /// <param term='commandName'>The name of the command to execute.</param>
        /// <param term='executeOption'>Describes how the command should be run.</param>
        /// <param term='varIn'>Parameters passed from the caller to the command handler.</param>
        /// <param term='varOut'>Parameters passed from the command handler to the caller.</param>
        /// <param term='handled'>Informs the caller if the command was handled or not.</param>
        /// <seealso class='Exec' />
        public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            handled = false;
            if (executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
            {
                if (commandName == "ILViewer.Connect.ILViewer")
                {
                    CreateIL();
                    handled = true;
                    return;
                }
            }
        }
        private DTE2 _applicationObject;
        private AddIn _addInInstance;

        #region custom Code

        private void CreateIL()
        {
            try
            {
                Paths.ActiveWindow = _applicationObject.ActiveWindow;
                string arguments;
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.ErrorDataReceived += (sender, e) => { File.WriteAllText(Paths.TempIL, "No IL"); };

                File.Delete(Paths.DllPath);
                File.Delete(Paths.TempIL);

                arguments = string.Format(@"/t:library /r:System.dll /out:""{0}"" ""{1}""", Paths.DllPath, Paths.ActiveFile);
                process.StartInfo = new ProcessStartInfo(Paths.Compilers.Current, arguments) { WindowStyle = ProcessWindowStyle.Hidden };
                process.Start();
                process.WaitForExit();

                File.WriteAllText(Paths.TempIL, "");

                arguments = string.Format("\"{0}\" /output=\"{1}\"", Paths.DllPath, Paths.TempIL);
                process.StartInfo = new ProcessStartInfo(Paths.ILdasmPath, arguments) { WindowStyle = ProcessWindowStyle.Hidden };
                process.Start();
                process.WaitForExit();
            }
            finally
            {
                if (!File.Exists(Paths.ViewIL) || File.ReadAllText(Paths.ViewIL).Trim() == "")
                {
                    File.WriteAllText(Paths.ViewIL, "No IL");
                }

                Window ilWindow = _applicationObject.ItemOperations.OpenFile(Paths.ViewIL, "");
                ilWindow.Document.ReplaceText(File.ReadAllText(Paths.ViewIL), File.ReadAllText(Paths.TempIL), 0);
                ilWindow.Document.Save(ilWindow.Document.FullName);
                Paths.ActiveWindow.SetFocus();

                string extension;
                foreach (Window window in _applicationObject.Windows)
                {
                    extension = Path.GetExtension(window.Document.FullName).ToLower();
                    if (extension == ".vb" || extension == ".cs")
                    {
                        _applicationObject.Events.get_DocumentEvents(window.Document).DocumentSaved += doc => { CreateIL(); };
                    }
                }
            }
        }

        #endregion
    }

    #region custom Code

    internal static class Paths
    {
        internal static string SDK;
        internal static string CurrentVSInstallFolder;
        internal static string FrameworkPath;
        internal static string ProgramFiles;
        internal static string ILdasmPath;
        internal static string Temp;
        private static Window activeWindow;

        internal static string ActiveFile;
        internal static string Extension;
        internal static string DllPath;
        internal static string TempIL;
        internal static string ViewIL;

        internal static Window ActiveWindow
        {
            get
            {
                return activeWindow;
            }
            set
            {
                activeWindow = value;
                ActiveDocument = ActiveWindow.Document;
                if (ActiveDocument != null)
                {
                    ActiveFile = ActiveDocument.FullName;
                    Extension = Path.GetExtension(ActiveFile).ToLower();
                    DllPath = Paths.Temp + "1.dll";
                    TempIL = ActiveFile.Substring(0, ActiveFile.LastIndexOf('.')) + "_temp" + Extension + ".il";
                    ViewIL = ActiveFile + ".il";
                }
                else
                {
                    ActiveFile = string.Empty;
                    Extension = string.Empty;
                    DllPath = string.Empty;
                    TempIL = string.Empty;
                    ViewIL = string.Empty;
                }
                if (Extension.Equals(".vb"))
                {
                    Compilers.Current = Compilers.VBC;
                }
                else if (Extension.Equals(".cs"))
                {
                    Compilers.Current = Compilers.CSC;
                }
                else
                {
                    Compilers.Current = string.Empty;
                }
            }
        }

        internal static Document ActiveDocument
        { get; set; }

        static Paths()
        {
            SDK = GetRegistry(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\.NETFramework", "InstallRoot");
            
            CurrentVSInstallFolder = GetRegistry(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SDKs\Windows", "CurrentInstallFolder");
            if (CurrentVSInstallFolder.Trim() == "")
                CurrentVSInstallFolder = GetRegistry(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SDKs\.NETFramework\v2.0", "InstallationFolder");
            FrameworkPath = SDK + @"\v2.0.50727\";
            ProgramFiles = Environment.GetEnvironmentVariable("programfile");
            ILdasmPath = CurrentVSInstallFolder + @"bin\ildasm.exe";
            Temp = Path.GetTempPath();
            Compilers.CSC = FrameworkPath + "csc.exe";
            Compilers.VBC = FrameworkPath + "vbc.exe";
        }

        private static string GetRegistry(string key, string node)
        {
            return Convert.ToString(Registry.GetValue(key, node, ""));
        }

        internal static class Compilers
        {
            internal static string CSC;
            internal static string VBC;
            internal static string Current;
        }
    }

    #endregion
}


