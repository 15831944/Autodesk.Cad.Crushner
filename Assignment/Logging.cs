using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Autodesk.Cad.Crushner.Assignment
{
    public class Logging : Core.Logging
    {
        private Logging() : base() { }

        protected override Core.Logging create() { return new Assignment.Logging() as Core.Logging; }

        private static String m_logFileName = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            , @"acad-2013-test.log");

        public static void AcEditorWriteMessage(string msg)
        {
            Database dbCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
            DocumentCollection acDocMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            Document acDoc = acDocMgr.GetDocument(dbCurrent);

            acDoc.Editor.WriteMessage(string.Format(@"{0}{1}", Environment.NewLine, msg));
        }

        public static void AcEditorWriteException(Exception e, string msg)
        {
            AcEditorWriteMessage(string.Format(@"{1}{0}{2}{0}{3}", Environment.NewLine, msg, e.Message, e.StackTrace));
        }

        public static void AcEditorDebugCaller(MethodBase methodBase)
        {
            DebugCaller(methodBase, string.Empty);
        }

        public static void AcEditorDebugCaller(MethodBase methodBase, string message)
        {
            Database dbCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
            DocumentCollection acDocMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            Document acDoc = acDocMgr.GetDocument(dbCurrent);

            String documentName = acDoc.Name;

            writeln(String.Format("{0}: Класс.Метод: {1}.{2}; Документ: {3}: {4}",
                DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss")
                , methodBase.ReflectedType.Name, methodBase.Name
                , documentName
                , message));
        }
    }
}
