using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;

using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace Autodesk.Cad.Crushner.Assignment
{
    public partial class MSExcel : Settings.MSExcel
    {
        public struct MAP_KEY_ENTITY
        {
            public Settings.MSExcel.COMMAND_ENTITY m_command;

            public Type m_type;

            public string m_nameSolidType;

            public string m_nameCreateMethod;
        }

        public static readonly List<MAP_KEY_ENTITY> s_MappingKeyEntity = new List<MAP_KEY_ENTITY>() {
            new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.CIRCLE, m_type = typeof(Circle), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.ARC, m_type = typeof(Arc), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.LINE, m_type = typeof(Line), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.PLINE3, m_type = typeof(Polyline3d), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.CONE, m_type = typeof(Solid3d), m_nameSolidType = @"Frustum", m_nameCreateMethod = @"CreateFrustum" }
            , new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.BOX, m_type = typeof(Solid3d), m_nameSolidType = @"Box", m_nameCreateMethod = @"CreateBox" }
            // линии - векторы
            , new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.ALINE_X, m_type = typeof(Line), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.ALINE_Y, m_type = typeof(Line), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.ALINE_Z, m_type = typeof(Line), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.RLINE_X, m_type = typeof(Line), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.RLINE_Y, m_type = typeof(Line), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.RLINE_Z, m_type = typeof(Line), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            // линия - сферическая СК
            , new MAP_KEY_ENTITY () { m_command = COMMAND_ENTITY.LINE_SPHERE, m_type = typeof(Line), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            ,
        };

        public static string GetNameSolidTypeEntity(COMMAND_ENTITY command)
        {
            string nameSolidTypeRes = string.Empty;

            MAP_KEY_ENTITY mapKeyEntity =
                s_MappingKeyEntity.Find(item => { return item.m_command == command; });

            if (!(mapKeyEntity.m_command == COMMAND_ENTITY.UNKNOWN))
                nameSolidTypeRes = mapKeyEntity.m_nameSolidType;
            else
                ;

            return nameSolidTypeRes;
        }

        public static string GetNameCreateMethodEntity(COMMAND_ENTITY command)
        {
            string nameMethodRes = string.Empty;

            MAP_KEY_ENTITY mapKeyEntity =
                s_MappingKeyEntity.Find(item => { return item.m_command == command; });

            if (!(mapKeyEntity.m_command == COMMAND_ENTITY.UNKNOWN))
                nameMethodRes = mapKeyEntity.m_nameCreateMethod;
            else
                ;

            return nameMethodRes;
        }

        public static Type GetTypeEntity(COMMAND_ENTITY command)
        {
            Type typeRes = Type.Missing as Type;

            MAP_KEY_ENTITY mapKeyEntity = 
                s_MappingKeyEntity.Find(item => { return item.m_command == command; });

            if (!(mapKeyEntity.m_command == COMMAND_ENTITY.UNKNOWN))
                typeRes = mapKeyEntity.m_type;
            else
                ;

            return typeRes;
        }

        public static MAP_KEY_ENTITY GetMapKeyEntity(COMMAND_ENTITY command)
        {
            MAP_KEY_ENTITY mapKeyEntityRes = new MAP_KEY_ENTITY() { m_command = COMMAND_ENTITY.UNKNOWN, m_type = Type.Missing as Type, m_nameCreateMethod = string.Empty, m_nameSolidType = string.Empty };

            MAP_KEY_ENTITY mapKeyEntity =
                s_MappingKeyEntity.Find(item => { return item.m_command == command; });

            return mapKeyEntityRes;
        }
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
        private delegate EntityCtor.ProxyEntity delegateNewEntity(Settings.EntityParser.ProxyEntity entity/*, string blockName*/);
        /// <summary>
        /// Тип делегата для упаковки примитива в строку таблицы
        ///  (подготовка к экспорту)
        /// </summary>
        /// <param name="pair">Сложный ключ - идентификатор примитива + объект примитива</param>
        /// <returns>Объект для последующего экспорта</returns>
        private delegate Settings.EntityParser.ProxyEntity delegateEntityToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair);
        /// <summary>
        /// Словарь с методами для создания/упаковки прмитивов разных типов
        ///  ключ - тип примитива
        /// </summary>
        private static Dictionary<COMMAND_ENTITY, METHODE_ENTITY> dictDelegateMethodeEntity = new Dictionary<COMMAND_ENTITY, METHODE_ENTITY>() {
            { COMMAND_ENTITY.CIRCLE, new METHODE_ENTITY () { newEntity = EntityCtor.newCircle, entityToDataRow = EntityCtor.circleToDataRow } }
            , { COMMAND_ENTITY.ARC, new METHODE_ENTITY () { newEntity = EntityCtor.newArc, entityToDataRow = EntityCtor.arcToDataRow } }
            , { COMMAND_ENTITY.LINE, new METHODE_ENTITY () { newEntity = EntityCtor.newLineDecart, entityToDataRow = EntityCtor.lineDecartToDataRow } }
            , { COMMAND_ENTITY.PLINE3, new METHODE_ENTITY () { newEntity = EntityCtor.newPolyLine3d, entityToDataRow = EntityCtor.polyLine3dToDataRow } }
            , { COMMAND_ENTITY.CONE, new METHODE_ENTITY () { newEntity = EntityCtor.newCone, entityToDataRow = EntityCtor.coneToDataRow } }
            , { COMMAND_ENTITY.BOX, new METHODE_ENTITY () { newEntity = EntityCtor.newBox, entityToDataRow = EntityCtor.boxToDataRow } }
            // векторы - линии
            , { COMMAND_ENTITY.ALINE_X, new METHODE_ENTITY () { newEntity = EntityCtor.newLineDecart, entityToDataRow = EntityCtor.alineDecartXToDataRow } }
            , { COMMAND_ENTITY.ALINE_Y, new METHODE_ENTITY () { newEntity = EntityCtor.newLineDecart, entityToDataRow = EntityCtor.alineDecartYToDataRow } }
            , { COMMAND_ENTITY.ALINE_Z, new METHODE_ENTITY () { newEntity = EntityCtor.newLineDecart, entityToDataRow = EntityCtor.alineDecartZToDataRow } }
            , { COMMAND_ENTITY.RLINE_X, new METHODE_ENTITY () { newEntity = EntityCtor.newLineDecart, entityToDataRow = EntityCtor.rlineDecartXToDataRow } }
            , { COMMAND_ENTITY.RLINE_Y, new METHODE_ENTITY () { newEntity = EntityCtor.newLineDecart, entityToDataRow = EntityCtor.rlineDecartYToDataRow } }
            , { COMMAND_ENTITY.RLINE_Z, new METHODE_ENTITY () { newEntity = EntityCtor.newLineDecart, entityToDataRow = EntityCtor.rlineDecartZToDataRow } }
             // линия - сферическая СК
            , { COMMAND_ENTITY.LINE_SPHERE, new METHODE_ENTITY () { newEntity = EntityCtor.newLineDecart, entityToDataRow = EntityCtor.lineSphereToDataRow } }
            ,
        };
        /// <summary>
        /// Словарь с определениями блоков
        /// </summary>
        public static MSExcel.DictionaryBlock s_dictBlock;
        /// <summary>
        /// Импортировать список объектов
        /// </summary>
        /// <param name="strNameSettingsExcelFile">Наименование файла конфигурации (книги MS Excel)</param>>
        /// /// <param name="format">Формат файла конфигурации (книги MS Excel)</param>>
        /// <returns>Признак ошибки при выполнении метода</returns>
        new public static int Import(string strNameSettingsExcelFile = @"", Settings.MSExcel.FORMAT format = Settings.MSExcel.FORMAT.ORDER)
        {
            int iErr = -1;

            try {
                Settings.MSExcel.Clear();
                iErr = Settings.MSExcel.Import(strNameSettingsExcelFile, format);

                Clear();
                Create(Settings.MSExcel.s_dictBlock);
            } catch (Exception e) {
                iErr = -1;

                Logging.AcEditorWriteException(e, strNameSettingsExcelFile);
            }

            return iErr;
        }

        public static EntityCtor.ProxyEntity GetEntityCtor(string blockName, KEY_ENTITY key)
        {
            return GetBlock(blockName).GetItem(key);
        }

        public static BLOCK GetBlock(string blockName)
        {
            return s_dictBlock[blockName] as BLOCK;
        }
        /// <summary>
        /// Тип делегата для определения типа примитива и его наименования
        /// </summary>
        /// <param name="ews">Лист книги MS Excel</param>
        /// <param name="rEntity">Строка таблицы со значениями параметров для примитива</param>
        /// <param name="comamand">Выходной параметр: тип примитива</param>
        /// <param name="name">Выходной параметр: наименование примитива</param>
        /// <returns>Результат выполнения делегата</returns>
        private delegate bool delegateTryParseEntity(Settings.EntityParser.ProxyEntity entity);

        new public static void Clear()
        {
            s_dictBlock?.Clear();
        }

        public static void Create(Settings.MSExcel.DictionaryBlock dictBlock)
        {
            if (s_dictBlock == null)
                s_dictBlock = new DictionaryBlock(dictBlock);
            else
                s_dictBlock.Create(dictBlock);
        }

        internal static void AddToExport(object entity)
        {
            throw new NotImplementedException();
        }
    }
}
