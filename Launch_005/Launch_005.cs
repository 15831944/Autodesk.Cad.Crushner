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

namespace Autodesk.Cad.Crushner.Launch_005
{
    public class Launch_005 : Core
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
        [CommandMethod("LAUNCH-005-IMPORT")]
        public void Import()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"LAUNCH-005-IMPORT"));

            reinitialize();

            MSExcel.Clear();
            clear();

            MSExcel.Import(@"settings_005.xls", MSExcel.FORMAT.HEAP);

            flash();
        }
        ///// <summary>
        ///// Метод плагина для выполнения команды - добавить прямую линию
        ///// </summary>
        //[CommandMethod("LAUNCH-005-EXPORT")]
        //public void Export()
        //{
        //    Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"LAUNCH-005-EXPORT"));

        //    reinitialize();

        //    export(@"settings_005.xls", MSExcel.FORMAT.HEAP);
        //}
        /// <summary>
        /// Метод плагина для выполнения команды - удалить все примитивы с чертежа
        /// </summary>
        [CommandMethod("LAUNCH-005-CLEAR_DRAWING")]
        public void ClearDrawing()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"LAUNCH-005-CLEAR_DRAWING"));

            reinitialize();

            clear();
        }
        /// <summary>
        /// Метод плагина для выполнения команды - очитстиь файл конфигурации
        /// </summary>
        [CommandMethod("LAUNCH-005-CLEAR_SETTINGS")]
        public void ClearSettings()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"LAUNCH-005-CLEAR_SETTINGS"));

            reinitialize();

            MSExcel.Clear();
        }
        /// <summary>
        /// Метод плагина для выполнения команды - отобразить все примитивы
        /// </summary>
        [CommandMethod("LAUNCH-005-PAINT")]
        public void Paint()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"LAUNCH-005-PAINT_ALL"));

            reinitialize();

            flash();
        }
        #endregion

        #region Управление (в т.ч. добавление) примитивов
        #endregion
    }
}
