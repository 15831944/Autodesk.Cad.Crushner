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

using Autodesk.Cad.Crushner.Assignment;

namespace Autodesk.Cad.Crushner.Command
{
    public class Command : Assignment.Command
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
        [CommandMethod("GO")]
        public void Go()
        {
            string command = @"GO";

            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполняется команда: {0} - перенаправление CRU-IMPORT", command));

            try {
                // !!! 2013
                Application.DocumentManager.MdiActiveDocument.SendStringToExecute(@"CRU-IMPORT ", true, false, true);
                // ??? 2014
                //Application.DocumentManager.MdiActiveDocument.Editor.Command(new object[] { @"CRU-IMPORT" });
            } catch (AutoCAD.Runtime.Exception e) {
                Logging.AcEditorWriteException(e, command);

                Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e);
            } catch (System.Exception e) {
                Logging.AcEditorWriteException(e, command);

                Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e);
            }
        }
        /// <summary>
        /// Метод плагина для выполнения команды - добавить окружность
        /// </summary>
        [CommandMethod("CRU-IMPORT")]
        public void Import()
        {
            string command = @"CRU-IMPORT";

            try {
                reinitialize();

                MSExcel.Clear();
                clear();

                MSExcel.Import(@"settings.xls", Settings.MSExcel.FORMAT.HEAP);

                flash();
            } catch (System.Exception e) {
                    Logging.AcEditorWriteException(e, command);

                    Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e);
            }

            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", command));
        }
        ///// <summary>
        ///// Метод плагина для выполнения команды - добавить прямую линию
        ///// </summary>
        //[CommandMethod("CRU-EXPORT")]
        //public void Export()
        //{
        //    Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"CRU-EXPORT"));

        //    reinitialize();

        //    export(@"settings.xls", MSExcel.FORMAT.HEAP);
        //}
        /// <summary>
        /// Метод плагина для выполнения команды - удалить все примитивы с чертежа
        /// </summary>
        [CommandMethod("CRU-CLEAR_DRAWING")]
        public void ClearDrawing()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"CRU-CLEAR_DRAWING"));

            reinitialize();

            clear();
        }
        /// <summary>
        /// Метод плагина для выполнения команды - очитстиь файл конфигурации
        /// </summary>
        [CommandMethod("CRU-CLEAR_SETTINGS")]
        public void ClearSettings()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"CRU-CLEAR_SETTINGS"));

            reinitialize();

            MSExcel.Clear();
        }
        /// <summary>
        /// Метод плагина для выполнения команды - отобразить все примитивы
        /// </summary>
        [CommandMethod("CRU-PAINT")]
        public void Paint()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"CRU-PAINT_ALL"));

            reinitialize();

            flash();
        }
        #endregion

        #region Управление (в т.ч. добавление) примитивов
        #endregion
    }
}
