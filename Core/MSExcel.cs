using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;

using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using static Autodesk.Cad.Crushner.Settings.Collection;

namespace Autodesk.Cad.Crushner.Core
{
    public partial class MSExcel : Collection
    {
        /// <summary>
        /// Структура для хранения методов распаковки и упаковки примитива из/в строки(у) таблицы
        /// </summary>
        private struct METHODE_ENTITY
        {
            public delegateNewEntity newEntity;

            public delegateEntityToDataRow entityToDataRow;
        }
        /// <summary>
        /// Тип делегата для создания примитива
        /// </summary>
        /// <param name="rEntity">Строка талицы со значениями для создания примитива</param>
        /// <param name="format">Формат книги MS Excel</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Созданный объект примитива</returns>
        private delegate EntityCtor.ProxyEntity delegateNewEntity(DataRow rEntity, FORMAT format/*, string blockName*/);
        /// <summary>
        /// Тип делегата для упаковки примитива в строку таблицы
        ///  (подготовка к экспорту)
        /// </summary>
        /// <param name="pair">Сложный ключ - идентификатор примитива + объект примитива</param>
        /// <param name="format">Формат книги MS Excel</param>
        /// <returns></returns>
        private delegate object[] delegateEntityToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair, FORMAT format);
        /// <summary>
        /// Словарь с методами для создания/упаковки прмитивов разных типов
        ///  ключ - тип примитива
        /// </summary>
        private static Dictionary<COMMAND_ENTITY, METHODE_ENTITY> dictDelegateMethodeEntity = new Dictionary<COMMAND_ENTITY, METHODE_ENTITY>() {
            { COMMAND_ENTITY.CIRCLE, new METHODE_ENTITY () { newEntity = EntityCtor.newCircle, entityToDataRow = EntityCtor.circleToDataRow } }
            , { COMMAND_ENTITY.ARC, new METHODE_ENTITY () { newEntity = EntityCtor.newArc, entityToDataRow = EntityCtor.arcToDataRow } }
            , { COMMAND_ENTITY.LINE, new METHODE_ENTITY () { newEntity = EntityCtor.newLine, entityToDataRow = EntityCtor.lineToDataRow } }
            , { COMMAND_ENTITY.PLINE3, new METHODE_ENTITY () { newEntity = EntityCtor.newPolyLine3d, entityToDataRow = EntityCtor.polyLine3dToDataRow } }
            , { COMMAND_ENTITY.CONE, new METHODE_ENTITY () { newEntity = EntityCtor.newCone, entityToDataRow = EntityCtor.coneToDataRow } }
            , { COMMAND_ENTITY.BOX, new METHODE_ENTITY () { newEntity = EntityCtor.newBox, entityToDataRow = EntityCtor.boxToDataRow } }
            // векторы - линии
            , { COMMAND_ENTITY.ALINE_X, new METHODE_ENTITY () { newEntity = EntityCtor.newALineX, entityToDataRow = EntityCtor.alineXToDataRow } }
            , { COMMAND_ENTITY.ALINE_Y, new METHODE_ENTITY () { newEntity = EntityCtor.newALineY, entityToDataRow = EntityCtor.alineYToDataRow } }
            , { COMMAND_ENTITY.ALINE_Z, new METHODE_ENTITY () { newEntity = EntityCtor.newALineZ, entityToDataRow = EntityCtor.alineZToDataRow } }
            , { COMMAND_ENTITY.RLINE_X, new METHODE_ENTITY () { newEntity = EntityCtor.newRLineX, entityToDataRow = EntityCtor.rlineXToDataRow } }
            , { COMMAND_ENTITY.RLINE_Y, new METHODE_ENTITY () { newEntity = EntityCtor.newRLineY, entityToDataRow = EntityCtor.rlineYToDataRow } }
            , { COMMAND_ENTITY.RLINE_Z, new METHODE_ENTITY () { newEntity = EntityCtor.newRLineZ, entityToDataRow = EntityCtor.rlineZToDataRow } }
            ,
        };
        /// <summary>
        /// Импортировать список объектов
        /// </summary>
        /// <returns>Список объектов</returns>
        public static int Import(string strNameSettingsExcelFile = @"", Settings.MSExcel.FORMAT format = Settings.MSExcel.FORMAT.ORDER)
        {
            int iErr = -1;

            

            try {
                iErr = Settings.MSExcel.Import();
            } catch (Exception e) {
                iErr = -1;

                Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e);

                Logging.AcEditorWriteException(e, strNameSettings);
            }

            return iErr;
        }
        /// <summary>
        /// Тип делегата для определения типа примитива и его наименования
        /// </summary>
        /// <param name="ews">Лист книги MS Excel</param>
        /// <param name="rEntity">Строка таблицы со значениями параметров для примитива</param>
        /// <param name="comamand">Выходной параметр: тип примитива</param>
        /// <param name="name">Выходной параметр: наименование примитива</param>
        /// <returns>Результат выполнения делегата</returns>
        private delegate bool delegateTryParseCommandAndNameEntity(ExcelWorksheet ews, DataRow rEntity, out COMMAND_ENTITY comamand, out string name);
        /// <summary>
        /// Словарь с делегатами для определения типа примитива и его наименования, ключ=формат файла
        /// </summary>
        private static Dictionary<FORMAT, delegateTryParseCommandAndNameEntity> dictDelegateTryParseCommandAndNameEntity = new Dictionary<FORMAT, delegateTryParseCommandAndNameEntity>() {
            { FORMAT.HEAP, tryParseCommandAndNameEntityHeap }
            , { FORMAT.ORDER, tryParseCommandAndNameEntityOrder }
        };
        /// <summary>
        /// Метод для определения типа примитива и его наименования в формате 'HEAP'
        /// </summary>
        /// <param name="ews">Лист книги MS Excel</param>
        /// <param name="rEntity">Строка таблицы со значениями параметров для примитива</param>
        /// <param name="command">Выходной параметр: тип примитива</param>
        /// <param name="name">Выходной параметр: наименование примитива</param>
        /// <returns>Результат выполнения делегата</returns>
        private static bool tryParseCommandAndNameEntityHeap(ExcelWorksheet ews, DataRow rEntity, out COMMAND_ENTITY command, out string name)
        {
            bool bRes = false;

            command = COMMAND_ENTITY.UNKNOWN;
            name = string.Empty;

            // 0-ой столбец - наименование
            // 1-ый столбец - тип примитива

            if (bRes = COMMAND_ENTITY.TryParse(rEntity[1].ToString(), true, out command) == true) {
                if (bRes = !(rEntity[0] is DBNull)) {
                    name = rEntity[0].ToString();

                    bRes = !(name.Equals(string.Empty) == true);
                } else
                    ;
            } else
                ;

            return bRes;
        }
        /// <summary>
        /// Метод для определения типа примитива и его наименования в формате 'ORDER'
        /// </summary>
        /// <param name="ews">Лист книги MS Excel</param>
        /// <param name="rEntity">Строка таблицы со значениями параметров для примитива</param>
        /// <param name="command">Выходной параметр: тип примитива</param>
        /// <param name="name">Выходной параметр: наименование примитива</param>
        /// <returns>Результат выполнения делегата</returns>
        private static bool tryParseCommandAndNameEntityOrder(ExcelWorksheet ews, DataRow rEntity, out COMMAND_ENTITY command, out string name)
        {
            bool bRes = false;

            command = COMMAND_ENTITY.UNKNOWN;
            name = string.Empty;

            if (bRes = COMMAND_ENTITY.TryParse(ews.Name, true, out command) == true) {
                if (bRes = !(rEntity[@"Name"] is DBNull)) {
                    name = rEntity[@"Name"].ToString();

                    bRes = !(name.Equals(string.Empty) == true);
                } else
                    ;
            } else
                ;

            return bRes;
        }

        internal static void AddToExport(object entity)
        {
            throw new NotImplementedException();
        }
    }
}
