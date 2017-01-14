using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
//using Autodesk.Windows;
//using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.IO;
using System.Reflection;

using Autodesk.Cad.Crushner.Common;

namespace Autodesk.Cad.Crushner.Launch_002
{
    public class Launch_002 : Launch_000
    {
        #region Обязательные методы плюгина
        /// <summary>
        /// Функция инициализации (выполняется при загрузке плагина)
        /// </summary>
        public override void Initialize()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod());

            base.Initialize();

            initialize();
        }

        protected override void reinitialize()
        {
            initialize();

            base.reinitialize();
        }

        private void initialize()
        {
            //!!! Волшебство. Без создания этого объекта ничего не работает. Должен быть вызов в собственной функции.
            Database dbCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
        }
        #endregion

        #region Обработчики событий
        #endregion

        #region Методы плагина для выполнения команд
        /// <summary>
        /// Метод плагина для выполнения команды - добавить окружность
        /// </summary>
        [CommandMethod("LAUNCH-002-IMPORT")]
        public void Import()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"LAUNCH-002-IMPORT"));

            reinitialize();

            MSExcel.Clear();
            clear();

            MSExcel.Import();

            flash();
        }
        /// <summary>
        /// Метод плагина для выполнения команды - добавить прямую линию
        /// </summary>
        [CommandMethod("LAUNCH-002-EXPORT")]
        public void Export()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"LAUNCH-002-EXPORT"));

            reinitialize();

            export(@"settings.xls", MSExcel.FORMAT.ORDER);
        }
        #endregion

        #region Управление (в т.ч. добавление) примитивов
        #endregion
    }
}
