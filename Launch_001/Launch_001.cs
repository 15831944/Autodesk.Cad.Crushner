using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
//using Autodesk.Windows;
//using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.IO;
using System.Reflection;

using Autodesk.Cad.Crushner.Common;

namespace Autodesk.Cad.Crushner.Launch_001
{
    public class Launch_001 : Launch_000 
    {
        /// <summary>
        /// Словарь акций-методов для создания того или иного типа примтивов
        /// </summary>
        private Dictionary<COMMAND_ENTITY, Action> _dictCommandEntityAction;
        /// <summary>
        /// Текущя команда на создание типа примитива
        /// </summary>
        private COMMAND_ENTITY _commandEntity;

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
            base.reinitialize();

            initialize();
        }

        private void initialize()
        {
            _dictCommandEntityAction = new Dictionary<COMMAND_ENTITY, Action>() {
                { COMMAND_ENTITY.CIRCLE, circleAdd }
                , { COMMAND_ENTITY.ARC, arcAdd }
                , { COMMAND_ENTITY.LINE, lineAdd }
                ,
            };

            _commandEntity = COMMAND_ENTITY.UNKNOWN;
        }
        #endregion

        #region Обработчики событий
        ///// <summary>
        ///// Обработчик события - создание нового документа завершено
        ///// </summary>
        ///// <param name="sender">Объект, инициировавший событие</param>
        ///// <param name="e">Аргумент события</param>
        //protected override void AcDocMgr_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        //{
        //    base.AcDocMgr_DocumentCreated(sender, e);

        //    if (!(_commandEntity == COMMAND_ENTITY.UNKNOWN))
        //        entityAction();
        //    else
        //        ;
        //}
        ///// <summary>
        ///// Обработчик события - создание нового документа завершено
        ///// </summary>
        ///// <param name="sender">Объект, инициировавший событие</param>
        ///// <param name="e">Аргумент события</param>
        //protected override void Application_Idle(object sender, EventArgs e)
        //{
        //    base.Application_Idle(sender, e);
        //}
        #endregion

        #region Методы плагина для выполнения команд
        /// <summary>
        /// Метод плагина для выполнения команды - добавить окружность
        /// </summary>
        [CommandMethod("LAUNCH-001-CIRCLE-ADD")]
        public void CircleAdd()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"LAUNCH-001-CIRCLE-ADD"));

            entityAdd(COMMAND_ENTITY.CIRCLE);
        }
        /// <summary>
        /// Метод плагина для выполнения команды - добавить прямую линию
        /// </summary>
        [CommandMethod("LAUNCH-001-LINE-ADD")]
        public void BoxAdd()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"LAUNCH-001-LINE-ADD"));

            entityAdd(COMMAND_ENTITY.LINE);
        }
        /// <summary>
        /// Метод плагина для выполнения команды - добавить дугу
        /// </summary>
        [CommandMethod("LAUNCH-001-ARC-ADD")]
        public void ArcAdd()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Выполнена команда: {0}", @"LAUNCH-001-ARC-ADD"));

            entityAdd(COMMAND_ENTITY.ARC);
        }
        #endregion

        #region Управление (в т.ч. добавление) примитивов
        /// <summary>
        /// Добавить окружность
        /// </summary>
        private void circleAdd ()
        {
            Database dbCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;

            BlockTableRecord btrCurrSpace;
            Circle cNewCircle = new Circle();
            ObjectId oidCircle;
            Vector3d norm;
            int cntEntity = Collection.GetCount(cNewCircle);

            using (Transaction trAdding = dbCurrent.TransactionManager.StartTransaction()) {
                cNewCircle.Center = new Point3d(25, 25, 25);
                cNewCircle.Radius = (50 * Scale) + ((5 * cntEntity) * Scale);
                cNewCircle.ColorIndex = 5 + cntEntity;
                switch (cntEntity % 3) {
                    case 0:
                        norm = Vector3d.XAxis;
                        break;
                    case 1:
                        norm = Vector3d.YAxis;
                        break;
                    default: // 3
                        norm = Vector3d.ZAxis;
                        break;
                }
                cNewCircle.TransformBy(Matrix3d.PlaneToWorld(norm));

                btrCurrSpace = trAdding.GetObject(dbCurrent.CurrentSpaceId
                    , OpenMode.ForWrite) as BlockTableRecord;

                Collection.Add(new EntityParser.ProxyEntity (cNewCircle));
                oidCircle = btrCurrSpace.AppendEntity(cNewCircle);
                trAdding.AddNewlyCreatedDBObject(cNewCircle, true);

                trAdding.Commit();
            }

            _commandEntity = COMMAND_ENTITY.UNKNOWN;

            DocumentCollection acDocMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            Document acDoc = acDocMgr.GetDocument(dbCurrent);
            acDoc.Editor.WriteMessage(string.Format(@"{0}Создан примитив {1}...", Environment.NewLine, typeof(Circle).Name));
        }

        private void arcAdd()
        {
            Database dbCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
            
            BlockTableRecord btrCurrSpace;
            Arc cNewArc = new Arc();
            ObjectId oidArc;
            Vector3d norm;
            int cntEntity = Collection.GetCount(cNewArc);

            using (Transaction trAdding = dbCurrent.TransactionManager.StartTransaction()) {
                cNewArc.Center = new Point3d(25, 25, 25);
                cNewArc.Radius = (50 * Scale) + ((5 * cntEntity) * Scale);
                cNewArc.ColorIndex = 5 + cntEntity;
                (cNewArc as Arc).StartAngle = 0.0F; (cNewArc as Arc).EndAngle = 1.57F;
                switch (cntEntity % 3) {
                    case 0:
                        norm = Vector3d.XAxis;
                        break;
                    case 1:
                        norm = Vector3d.YAxis;
                        break;
                    default: // 3
                        norm = Vector3d.ZAxis;
                        break;
                }
                cNewArc.TransformBy(Matrix3d.PlaneToWorld(norm));

                btrCurrSpace = trAdding.GetObject(dbCurrent.CurrentSpaceId
                    , OpenMode.ForWrite) as BlockTableRecord;

                Collection.Add(new EntityParser.ProxyEntity (cNewArc));
                oidArc = btrCurrSpace.AppendEntity(cNewArc);
                trAdding.AddNewlyCreatedDBObject(cNewArc, true);

                trAdding.Commit();
            }

            _commandEntity = COMMAND_ENTITY.UNKNOWN;

            DocumentCollection acDocMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            Document acDoc = acDocMgr.GetDocument(dbCurrent);
            acDoc.Editor.WriteMessage(string.Format(@"{0}Создан примитив {1}...", Environment.NewLine, typeof(Arc).Name));
        }

        private void lineAdd()
        {
            Database dbCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
            DocumentCollection acDocMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            Document acDoc = acDocMgr.GetDocument(dbCurrent);
            acDoc.Editor.WriteMessage(string.Format(@"{0}Тип примитива {1} в текущей версии недоступен...", Environment.NewLine, typeof(Line).Name));
        }
        #endregion

        #region Общие методы по работе с примитивами
        /// <summary>
        /// Добавить примитив в соответсвии с аргументом
        /// </summary>
        /// <param name="commandEntity">Тип примитива</param>
        private void entityAdd(COMMAND_ENTITY commandEntity)
        {
            reinitialize();

            _commandEntity = commandEntity;

            entityAction();
        }
        /// <summary>
        /// Выполнить действие по добавлению примитива в чертеж
        /// </summary>
        private void entityAction()
        {
            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"_commandEntity={0}", _commandEntity.ToString()));

            if (_dictCommandEntityAction.ContainsKey(_commandEntity) == true)
                _dictCommandEntityAction[_commandEntity]();
            else {
                Database dbCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
                DocumentCollection acDocMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                Document acDoc = acDocMgr.GetDocument(dbCurrent);
                acDoc.Editor.WriteMessage(string.Format(@"{0}Неизвестная команда создания притмитива {1}...", Environment.NewLine, _commandEntity.ToString()));
            }
        }
        #endregion
    }
}
