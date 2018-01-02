using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Autodesk.Cad.Crushner.Core
{
    public class ProgramBase
    {
        public static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}
