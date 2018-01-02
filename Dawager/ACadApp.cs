
using Autodesk.Cad.Crushner.Core;
using System;
using System.Reflection;
using System.Runtime.InteropServices;

//using Autodesk.AutoCAD.ApplicationServices;
//using Autodesk.AutoCAD.ApplicationServices.Core;
//using Autodesk.AutoCAD.;
//using Autodesk.AutoCAD.Runtime;
//using Autodesk.AutoCAD.ApplicationServices;

namespace Dawager
{
    public class ACadApp
    {
        private static
            dynamic
            //Application
            _application;

        static string _autocadClassId = "AutoCAD.Application";

        private static void GetAutoCAD()
        {
            _application = Marshal.GetActiveObject(_autocadClassId)
                //as Application
                ;
        }

        private static void startAutoCad()
        {
            var t = Type.GetTypeFromProgID(_autocadClassId, true);
            // Create a new instance Autocad.
            var obj = Activator.CreateInstance(t, true);
            // No need for casting with dynamics
            _application = obj;

            ////(
            //    _application
            ////as Application)
            //        .Visible = true;
        }

        private static void stopAutoCad()
        {
            if (!(_application == null)) {
                //_application.Visible = false;

                _application.Quit();

                _application = null;
            } else
                ;
        }

        public static bool EnsureAutoCadIsRunning(string classId, bool bVerify)
        {
            //// вариант №1
            //string AutoCAD = GetAutoCADLaunchPath();
            //if (AutoCAD.Length != null || AutoCAD.Length > 0)
            //    System.Diagnostics.Process.Start(AutoCAD);
            //else
            //    ;

            // вариант №2
            if (!string.IsNullOrEmpty(classId) && classId != _autocadClassId)
                _autocadClassId = classId;
            else
                ;

            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Loading Autocad: {0}", _autocadClassId));

            if (object.Equals(_application, null) == true) {
                try {
                    GetAutoCAD();
                } catch (COMException ex) {
                    try {
                        if (bVerify == false) {
                            startAutoCad();
                        } else
                            ;
                    } catch (Exception e2x) {
                        Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e2x);

                        ThrowComException(ex);
                    }
                } catch (Exception ex) {
                    ThrowComException(ex);
                }
            } else
                if (bVerify == false)
                    stopAutoCad();
                else
                    ;

            return !(_application == null);
        }

        private static void ThrowComException(Exception ex)
        {
            Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), ex);
        }

        public static string GetAutoCADLaunchPath()
        {
            Microsoft.Win32.RegistryKey RegistryKeyRoot =
                   Microsoft.Win32.Registry.CurrentUser;
            if (RegistryKeyRoot != null)
            {
                Microsoft.Win32.RegistryKey RegistryKeySoftware =
                    RegistryKeyRoot.OpenSubKey("Software");

                if (RegistryKeySoftware != null)
                {
                    Microsoft.Win32.RegistryKey RegistryKeyAutodesk =
                           RegistryKeySoftware.OpenSubKey("Autodesk");

                    if (RegistryKeyAutodesk != null)
                    {
                        Microsoft.Win32.RegistryKey RegistryKeyDWGCommon =
                      RegistryKeyAutodesk.OpenSubKey("DWGCommon");

                        if (RegistryKeyDWGCommon != null)
                        {
                            Microsoft.Win32.RegistryKey RegistryKeyApps =
                       RegistryKeyDWGCommon.OpenSubKey("shellex").OpenSubKey("Apps");

                            if (RegistryKeyApps != null)
                            {
                                string GUID = RegistryKeyApps.GetValue("").ToString();
                                string Path = RegistryKeyApps.OpenSubKey(GUID).GetValue("OpenLaunch").ToString();

                                Path = Path.Replace("\"%1\"", "");
                                Path = Path.Replace("\"", "");
                                Path = Path.Trim();

                                return Path;
                            }
                        }
                    }
                }
            }

            return "";
        }

        public static void NewDocument()
        {
            //// Display the application and return the name and version
            //(_application as AcadApplication).Visible = true;
            //System.Windows.Forms.MessageBox.Show("Now running " + acAppComObj.Name +
            //                                     " version " + acAppComObj.Version);

            //// Get the active document
            //AcadDocument acDocComObj;
            //acDocComObj = acAppComObj.ActiveDocument;

            //// Optionally, load your assembly and start your command or if your assembly
            //// is demandloaded, simply start the command of your in-process assembly.
            //acDocComObj.SendCommand("(command " + (char)34 + "NETLOAD" + (char)34 + " " +
            //                        (char)34 + @"C:\Users\Administrator\Documents\All Code\main-libraries\IOAutoCADHandler\bin\Debug\IOAutoCADHandler.dll" + (char)34 + ") ");

            //acDocComObj.SendCommand("DRAWCOMPONENT");
        }
    }
}