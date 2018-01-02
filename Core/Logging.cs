using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Autodesk.Cad.Crushner.Core
{
    public class Logging
    {
        protected Logging()
        {
            if (_this == null)
                _this = create();
            else
                ;
        }

        protected virtual Logging create () {
            return new Logging();
        }

        protected Logging _this;
        /// <summary>
        /// Объект класса логгирования для доступа к нестатическим
        /// </summary>
        public Logging This { get { return _this; } }

        private static String m_logFileName = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            , @"acad-2013-test.log");

        public static void ExceptionCaller(MethodBase methodBase, System.Exception e)
        {
            ExceptionCaller(methodBase, e, string.Empty);
        }

        public static void ExceptionCaller(MethodBase methodBase, System.Exception e, string message)
        {
            writeln(String.Format("{1}: Класс.Метод: {2}.{3}{0}{4}{0}{5}{0}{6}",
                Environment.NewLine
                , DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss")
                , methodBase.ReflectedType.Name, methodBase.Name
                , message
                , e.Message
                , e.StackTrace));
        }

        public static void DebugCaller(MethodBase methodBase)
        {
            DebugCaller (methodBase, string.Empty);
        }

        public static void DebugCaller(MethodBase methodBase, string message)
        {
            writeln(String.Format("{0}: Класс.Метод: {1}.{2} - {3}",
                DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss")
                , methodBase.ReflectedType.Name, methodBase.Name
                , message));
        }

        protected static void writeln(string line)
        {
            using (StreamWriter sw =
                File.AppendText(m_logFileName))
            {
                sw.WriteLine(line);
                sw.Flush();
                sw.Close();
            }
        }
    }
}
