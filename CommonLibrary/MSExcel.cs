using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

using GemBox.Spreadsheet;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using System.Reflection;
using System.IO;

using excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Autodesk.Cad.Crushner.Common
{
    /// <summary>
    /// Перечисление - типы примитивов
    /// </summary>
    public enum COMMAND_ENTITY : short
    {
        UNKNOWN = -1
        , CIRCLE, ARC, LINE
            , COUNT
    }
    /// <summary>
    /// Уникалбный идентификатор примитива на чертеже
    /// </summary>
    public struct KEY_ENTITY
    {
        private static Char s_chNameDelimeter = '-';

        public COMMAND_ENTITY m_command;

        public int m_index;
        /// <summary>
        /// Уникальное наименование со специальной сигнатурой (состоит из 2-х частей)
        ///  1 - COMMAND_ENTITY
        ///  2 - 3-х значный цифровой индекс
        /// </summary>
        public string Name {
            get {
                if (!(m_command == COMMAND_ENTITY.UNKNOWN))
                    if (!(m_index < 0))
                        return string.Format(@"{0}{2}{1:000}", m_command.ToString(), m_index, s_chNameDelimeter);
                    else
                        return string.Format(@"{0}", m_command.ToString());
                else
                    return string.Empty;
            }
        }
        /// <summary>
        /// Тип объекта
        /// </summary>
        public Type m_type;

        public KEY_ENTITY(string name)
        {
            string[] names = name.Split(s_chNameDelimeter);

            if ((Enum.TryParse(names[0], out m_command) == true)
                && (Int32.TryParse(names[1], out m_index) == true))
                m_type = MSExcel.GetTypeEntity(m_command);
            else {
                m_command = COMMAND_ENTITY.UNKNOWN;
                m_index = -1;
                m_type = Type.Missing as Type;
            }
        }

        public KEY_ENTITY(Type type, COMMAND_ENTITY command, int indx)
        {
            m_type = type;

            m_command = command;

            m_index = indx;
        }

        public static bool operator==(KEY_ENTITY o1, KEY_ENTITY o2)
        {
            return (o1.m_command.Equals(o2.m_command) == true)
                && (o1.m_index.Equals(o2.m_index) == true)
                && (o1.m_type.Equals(o2.m_type) == true);
        }

        public static bool operator !=(KEY_ENTITY o1, KEY_ENTITY o2)
        {
            return (o1.m_command.Equals(o2.m_command) == false)
                || (o1.m_index.Equals(o2.m_index) == false)
                || (o1.m_type.Equals(o2.m_type) == false);
        }

        public override bool Equals(object obj)
        {
            return this == (KEY_ENTITY)obj;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }

    public class Collection
    {
        protected static readonly List<KEY_ENTITY> s_MappingKeyEntity = new List<KEY_ENTITY>() {
            new KEY_ENTITY () { m_command = COMMAND_ENTITY.CIRCLE, m_index = -1, m_type = typeof(Circle) }
            , new KEY_ENTITY () { m_command = COMMAND_ENTITY.ARC, m_index = -1, m_type = typeof(Arc) }
            , new KEY_ENTITY () { m_command = COMMAND_ENTITY.LINE, m_index = -1, m_type = typeof(Line) }
            ,
        };

        public static Type GetTypeEntity(COMMAND_ENTITY command)
        {
            Type typeRes = Type.Missing as Type;

            KEY_ENTITY keyEntity = 
                s_MappingKeyEntity.Find(item => { return item.m_command == command; });

            if (!(keyEntity == null))
                typeRes = keyEntity.m_type;
            else
                ;

            return typeRes;
        }

        public static Dictionary<KEY_ENTITY, Entity> s_dictEntity = new Dictionary<KEY_ENTITY, Entity>();

        public static void Clear()
        {
            s_dictEntity.Clear();
        }

        public static void Add(Entity entity)
        {
            s_dictEntity.Add(GetKeyEntity(entity.GetType()), entity);
        }

        public static void Add(string name, Entity entity)
        {
            s_dictEntity.Add(new KEY_ENTITY(name), entity);
        }

        public static KEY_ENTITY GetKeyEntity(Type typeEntity)
        {
            return new KEY_ENTITY(typeEntity
                , s_MappingKeyEntity.Find(type => type.m_type == typeEntity).m_command
                , GetCount(typeEntity));
        }

        public static int GetCount(Type typeEntity)
        {
            int iRes = 0;

            iRes = s_dictEntity.Keys.Where(type => type.m_type == typeEntity).Count();

            return iRes;
        }
    }

    public class MSExcel : Collection
    {
        public enum FORMAT : short { UNKNOWN = -1, ORDER, HEAP }
        /// <summary>
        /// Наименование книги MS Excel со списком объектов по умолчанию
        /// </summary>
        private static string s_nameSettings = @"settings.xls";

        private static string getFullNameSettingsExcelFile(string strNameSettingsExcelFile = @"")
        {
            return string.Format(@"{0}\{1}"
                //, AppDomain.CurrentDomain.BaseDirectory
                , Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                , strNameSettingsExcelFile.Equals(string.Empty) == false ?
                    strNameSettingsExcelFile :
                        s_nameSettings);
        }

        private static Dictionary<string, System.Data.DataTable> _dictDataTableOfExcelWorksheet;
        /// <summary>
        /// Создать таблицу для проецирования значений с листа книги MS Excel
        ///  , где наименования полей таблицы содержатся в 0-ой строке листа книги MS Excel
        /// </summary>
        /// <param name="nameWorksheet">Наименование листа книги MS Excel</param>
        /// <param name="rg">Регион на листе(странице) книги MS Excel</param>
        /// <param name="format">Формат книги MS Excel</param>
        private static void createDataTableWorksheet (string nameWorksheet
            , GemBox.Spreadsheet.CellRange rg
            , FORMAT format = FORMAT.ORDER) {
            string nameColumn = string.Empty;

            if (_dictDataTableOfExcelWorksheet == null)
            // добавить элемент
                _dictDataTableOfExcelWorksheet = new Dictionary<string, System.Data.DataTable>();
            else
                ;

            if (_dictDataTableOfExcelWorksheet.Keys.Contains(nameWorksheet) == false)
            // добавить таблицу (пустую)
                _dictDataTableOfExcelWorksheet.Add(nameWorksheet, new System.Data.DataTable());
            else
                ;

            if (!(_dictDataTableOfExcelWorksheet[nameWorksheet].Columns.Count == (rg.LastColumnIndex + 1))) {
                _dictDataTableOfExcelWorksheet[nameWorksheet].Columns.Clear();

                for (int j = rg.FirstColumnIndex; !(j > rg.LastColumnIndex); j++) {
                   // наименование столбцов таблицы всегда в 0-ой строке листа книги MS Excel
                   switch (format) {
                        case FORMAT.HEAP:
                            nameColumn = string.Format(@"{0:000}", j);
                            break;
                        case FORMAT.ORDER:
                        default:
                            nameColumn = rg[0, j - rg.FirstColumnIndex].Value.ToString();
                            break;
                    }
                    // все поля в таблице типа 'string'
                    _dictDataTableOfExcelWorksheet[nameWorksheet].Columns.Add(nameColumn, typeof(string));
                }
            } else
                ;
        }
        /// <summary>
        /// Импортировать список объектов
        /// </summary>
        /// <returns>Список объектов</returns>
        public static int Import(string strNameSettingsExcelFile = @"", FORMAT format = FORMAT.ORDER)
        {
            int iErr = -1;

            string strNameSettings = getFullNameSettingsExcelFile(strNameSettingsExcelFile);
            //DataRow dataRow; // для самостоятельного заполнения таблицы
            //VALIDATE_CELL_RESULT iValidateCellRes = VALIDATE_CELL_RESULT.BREAK; // для самостоятельного заполнения таблицы
            int i = -1, j = -1;
            string cellValue = string.Empty;

            _dictDataTableOfExcelWorksheet = new Dictionary<string, System.Data.DataTable>();

            try {
                ExcelFile ef = new ExcelFile();
                ef.LoadXls(strNameSettings, XlsOptions.None);

                Logging.AcEditorWriteMessage(string.Format(@"Книга открыта, листов = {0}", ef.Worksheets.Count));

                iErr = import(ef, format);
            } catch (Exception e) {
                iErr = -1;

                Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e);

                Logging.AcEditorWriteException(e, strNameSettings);
            }

            return iErr;
        }

        private static int import(ExcelFile ef, FORMAT format)
        {
            int iRes = 0;

            Entity entity = null;
            COMMAND_ENTITY commandEntity = COMMAND_ENTITY.UNKNOWN;
            string nameEntity = string.Empty;

            foreach (ExcelWorksheet ews in ef.Worksheets) {
                GemBox.Spreadsheet.CellRange range = ews.GetUsedCellRange();

                Logging.AcEditorWriteMessage(string.Format(@"Обработка листа с имененм = {0}", ews.Name));

                if (range.LastRowIndex > 0) {
                    extractDataWorksheet(ews, range, format);

                    foreach (DataRow rEntity in _dictDataTableOfExcelWorksheet[ews.Name].Rows) {
                        if (dictDelegateTryParseCommandAndNameEntity[format](ews, rEntity, out commandEntity, out nameEntity) == true) {
                            entity = null;

                            // соэдать примитив 
                            if (dictDelegateMethodeEntity.ContainsKey(commandEntity) == true)
                                entity = dictDelegateMethodeEntity[commandEntity].newEntity(rEntity, format);
                            else
                                ;

                            if (!(entity == null))
                                Add(
                                    nameEntity
                                    , entity
                                );
                            else
                                Logging.AcEditorWriteMessage(string.Format(@"Элемент с именем {0} пропущен..."
                                    , nameEntity
                                ));
                        } else
                        // ошибка при получении типа и наименования примитива
                            ;
                    } // цикл по строкам таблицы для листа книги MS Excel

                    Logging.AcEditorWriteMessage(string.Format(@"На листе с имененм = {0} обработано строк = {1}, добавлено элементов {2}"
                        , ews.Name
                        , range.LastRowIndex
                        , _dictDataTableOfExcelWorksheet[ews.Name].Rows.Count
                    ));
                } else
                // нет строк с данными
                    Logging.AcEditorWriteMessage(string.Format(@"На листе с имененм = {0} нет строк для распознования", ews.Name));
            }

            return iRes;
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

        private struct METHODE_ENTITY
        {
            public delegateNewEntity newEntity;

            public delegateEntityToDataRow entityToDataRow;
        }

        private delegate Entity delegateNewEntity (DataRow rEntity, FORMAT format);

        private delegate object[] delegateEntityToDataRow(KeyValuePair<KEY_ENTITY, Entity> pair, FORMAT format);

        private static Dictionary<COMMAND_ENTITY, METHODE_ENTITY> dictDelegateMethodeEntity = new Dictionary<COMMAND_ENTITY, METHODE_ENTITY>() {
            { COMMAND_ENTITY.CIRCLE, new METHODE_ENTITY () { newEntity = newCircle, entityToDataRow = circleToDataRow } }
            , { COMMAND_ENTITY.ARC, new METHODE_ENTITY () { newEntity = newArc, entityToDataRow = arcToDataRow } }
            , { COMMAND_ENTITY.LINE, new METHODE_ENTITY () { newEntity = newLine, entityToDataRow = lineToDataRow } }
        };
        /// <summary>
        /// Перечисление - индексы столбцов на листе книги MS Excel в формате 'HEAP'
        /// </summary>
        private enum HEAP_INDEX_COLUMN {
            CIRCLE_CENTER_X = 2, CIRCLE_CENTER_Y, CIRCLE_CENTER_Z, CIRCLE_RADIUS, CIRCLE_COLORINDEX, CIRCLE_TICKNESS
            , ARC_CENTER_X = 2, ARC_CENTER_Y, ARC_CENTER_Z, ARC_RADIUS, ARC_COLORINDEX, ARC_TICKNESS, ARC_ANGLE_START, ARC_ANGLE_END
            , LINE_START_X = 2, LINE_START_Y, LINE_START_Z, LINE_END_X, LINE_END_Y, LINE_END_Z, LINE_COLORINDEX, LINE_TICKNESS
        }
        /// <summary>
        /// Создать новый примитив - окружность по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы</param>
        /// <param name="format">Фориат файла конфигурации из которого была импортирована таблица</param>
        /// <returns>Объект примитива - окружность</returns>
        private static Entity newCircle(DataRow rEntity, FORMAT format)
        {
            Entity entityRes = null;
            // соэдать примитив 
            entityRes = new Circle();
            // значения для параметров примитива
            switch (format) {
                case FORMAT.HEAP:
                    (entityRes as Circle).Center = new Point3d(
                        double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.CIRCLE_CENTER_X].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.CIRCLE_CENTER_Y].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.CIRCLE_CENTER_Z].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (entityRes as Circle).Radius = double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.CIRCLE_RADIUS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (entityRes as Circle).ColorIndex = int.Parse(rEntity[(int)HEAP_INDEX_COLUMN.CIRCLE_COLORINDEX].ToString());
                    (entityRes as Circle).Thickness = double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.CIRCLE_TICKNESS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                case FORMAT.ORDER:
                    (entityRes as Circle).Center = new Point3d(
                        double.Parse(rEntity[@"CENTER.X"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"CENTER.Y"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"CENTER.Z"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (entityRes as Circle).Radius = double.Parse(rEntity[@"Radius"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (entityRes as Circle).ColorIndex = int.Parse(rEntity[@"ColorIndex"].ToString());
                    (entityRes as Circle).Thickness = double.Parse(rEntity[@"TICKNESS"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                default:
                    break;
            }

            return entityRes;
        }
        /// <summary>
        /// Создать новый примитив - дугу по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы</param>
        /// <param name="format">Фориат файла конфигурации из которого была импортирована таблица</param>
        /// <returns>Объект примитива - дуга</returns>
        private static Entity newArc(DataRow rEntity, FORMAT format)
        {
            Entity entityRes = null;
            // соэдать примитив 
            entityRes = new Arc();
            // значения для параметров примитива
            switch (format) {
                case FORMAT.HEAP:
                    (entityRes as Arc).Center = new Point3d(
                        double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.ARC_CENTER_X].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.ARC_CENTER_Y].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.ARC_CENTER_Z].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (entityRes as Arc).Radius = double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.ARC_RADIUS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (entityRes as Arc).ColorIndex = int.Parse(rEntity[(int)HEAP_INDEX_COLUMN.ARC_COLORINDEX].ToString());
                    (entityRes as Arc).Thickness = double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.ARC_TICKNESS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (entityRes as Arc).StartAngle = (Math.PI / 180) * float.Parse(rEntity[(int)HEAP_INDEX_COLUMN.ARC_ANGLE_START].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (entityRes as Arc).EndAngle = (Math.PI / 180) * float.Parse(rEntity[(int)HEAP_INDEX_COLUMN.ARC_ANGLE_END].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                case FORMAT.ORDER:
                    (entityRes as Arc).Center = new Point3d(
                        double.Parse(rEntity[@"CENTER.X"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"CENTER.Y"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"CENTER.Z"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (entityRes as Arc).Radius = double.Parse(rEntity[@"Radius"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (entityRes as Arc).ColorIndex = int.Parse(rEntity[@"ColorIndex"].ToString());
                    (entityRes as Arc).Thickness = double.Parse(rEntity[@"TICKNESS"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (entityRes as Arc).StartAngle = (Math.PI / 180) * float.Parse(rEntity[@"ANGLE.START"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (entityRes as Arc).EndAngle = (Math.PI / 180) * float.Parse(rEntity[@"ANGLE.END"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                default:
                    break;
            }

            return entityRes;
        }
        /// <summary>
        /// Создать новый примитив - линия по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы</param>
        /// <param name="format">Фориат файла конфигурации из которого была импортирована таблица</param>
        /// <returns>Объект примитива - линия</returns>
        private static Entity newLine(DataRow rEntity, FORMAT format)
        {
            Entity entityRes = null;
            // соэдать примитив 
            entityRes = new Line();
            // значения для параметров примитива
            switch (format) {
                case FORMAT.HEAP:
                    (entityRes as Line).StartPoint = new Point3d(
                        double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.LINE_START_X].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.LINE_START_Y].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.LINE_START_Z].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (entityRes as Line).EndPoint = new Point3d(
                        double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.LINE_END_X].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.LINE_END_Y].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.LINE_END_Z].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (entityRes as Line).ColorIndex = int.Parse(rEntity[(int)HEAP_INDEX_COLUMN.LINE_COLORINDEX].ToString());
                    (entityRes as Line).Thickness = double.Parse(rEntity[(int)HEAP_INDEX_COLUMN.LINE_TICKNESS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                case FORMAT.ORDER:
                    (entityRes as Line).StartPoint = new Point3d(
                        double.Parse(rEntity[@"START.X"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"START.Y"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"START.Z"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (entityRes as Line).EndPoint = new Point3d(
                        double.Parse(rEntity[@"END.X"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"END.Y"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"END.Z"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (entityRes as Line).ColorIndex = int.Parse(rEntity[@"ColorIndex"].ToString());
                    (entityRes as Line).Thickness = double.Parse(rEntity[@"TICKNESS"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                default:
                    break;
            }

            return entityRes;
        }

        private static object[] circleToDataRow(KeyValuePair<KEY_ENTITY, Entity> pair, FORMAT format)
        {
            object[] rowRes = null;

            switch (format) {
                case FORMAT.HEAP:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0}", pair.Key.m_command.ToString()) //!!! COMMAND_ENTITY
                        , string.Format(@"{0:0.0}", (pair.Value as Circle).Center.X) //CENTER.X
                        , string.Format(@"{0:0.0}", (pair.Value as Circle).Center.Y) //CENTER.Y
                        , string.Format(@"{0:0.0}", (pair.Value as Circle).Center.Z) //CENTER.Z
                        , string.Format(@"{0:0}", (pair.Value as Circle).Radius) //Radius
                        , string.Format(@"{0}", (pair.Value as Circle).ColorIndex) //ColorIndex
                        , string.Format(@"{0:0.000}", (pair.Value as Circle).Thickness) //Tickness
                    };
                    break;
                case FORMAT.ORDER:
                default:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0:0.0}", (pair.Value as Circle).Center.X) //CENTER.X
                        , string.Format(@"{0:0.0}", (pair.Value as Circle).Center.Y) //CENTER.Y
                        , string.Format(@"{0:0.0}", (pair.Value as Circle).Center.Z) //CENTER.Z
                        , string.Format(@"{0:0}", (pair.Value as Circle).Radius) //Radius
                        , string.Format(@"{0}", (pair.Value as Circle).ColorIndex) //ColorIndex
                        , string.Format(@"{0:0.000}", (pair.Value as Circle).Thickness) //Tickness
                    };
                    break;
            }

            return rowRes;
        }

        private static object[] arcToDataRow(KeyValuePair<KEY_ENTITY, Entity> pair, FORMAT format)
        {
            object[] rowRes = null;

            switch (format) {
                case FORMAT.HEAP:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0}", pair.Key.m_command.ToString()) //!!! COMMAND_ENTITY
                        , string.Format(@"{0:0.0}", (pair.Value as Arc).Center.X) //CENTER.X
                        , string.Format(@"{0:0.0}", (pair.Value as Arc).Center.Y) //CENTER.Y
                        , string.Format(@"{0:0.0}", (pair.Value as Arc).Center.Z) //CENTER.Z
                        , string.Format(@"{0:0}", (pair.Value as Arc).Radius) //Radius
                        , string.Format(@"{0}", (pair.Value as Arc).ColorIndex) //ColorIndex
                        , string.Format(@"{0:0.000}", (pair.Value as Arc).Thickness) //Tickness
                        , string.Format(@"{0:0}", (180 / Math.PI) * (pair.Value as Arc).StartAngle) //START.ANGLE
                        , string.Format(@"{0:0}", (180 / Math.PI) * (pair.Value as Arc).EndAngle) //END.ANGLE
                    };
                    break;
                case FORMAT.ORDER:
                default:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0:0.0}", (pair.Value as Arc).Center.X) //CENTER.X
                        , string.Format(@"{0:0.0}", (pair.Value as Arc).Center.Y) //CENTER.Y
                        , string.Format(@"{0:0.0}", (pair.Value as Arc).Center.Z) //CENTER.Z
                        , string.Format(@"{0:0}", (pair.Value as Arc).Radius) //Radius
                        , string.Format(@"{0}", (pair.Value as Arc).ColorIndex) //ColorIndex
                        , string.Format(@"{0:0.000}", (pair.Value as Arc).Thickness) //Tickness
                        , string.Format(@"{0:0}", (180 / Math.PI) * (pair.Value as Arc).StartAngle) //START.ANGLE
                        , string.Format(@"{0:0}", (180 / Math.PI) * (pair.Value as Arc).EndAngle) //END.ANGLE
                    };
                    break;
            }

            return rowRes;
        }

        private static object[] lineToDataRow(KeyValuePair<KEY_ENTITY, Entity> pair, FORMAT format)
        {
            object[] rowRes = null;

            switch (format) {
                case FORMAT.HEAP:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0}", pair.Key.m_command.ToString()) //!!! COMMAND_ENTITY
                        , string.Format(@"{0:0.0}", (pair.Value as Line).StartPoint.X) //SATRT.X
                        , string.Format(@"{0:0.0}", (pair.Value as Line).StartPoint.Y) //START.Y
                        , string.Format(@"{0:0.0}", (pair.Value as Line).StartPoint.Z) //START.Z
                        , string.Format(@"{0:0.0}", (pair.Value as Line).EndPoint.X) //END.X
                        , string.Format(@"{0:0.0}", (pair.Value as Line).EndPoint.Y) //END.Y
                        , string.Format(@"{0:0.0}", (pair.Value as Line).EndPoint.Z) //END.Z
                        , string.Format(@"{0}", (pair.Value as Line).ColorIndex) //ColorIndex
                        , string.Format(@"{0:0.000}", (pair.Value as Line).Thickness) //Tickness
                    };
                    break;
                case FORMAT.ORDER:
                default:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0:0.0}", (pair.Value as Line).StartPoint.X) //SATRT.X
                        , string.Format(@"{0:0.0}", (pair.Value as Line).StartPoint.Y) //START.Y
                        , string.Format(@"{0:0.0}", (pair.Value as Line).StartPoint.Z) //START.Z
                        , string.Format(@"{0:0.0}", (pair.Value as Line).EndPoint.X) //END.X
                        , string.Format(@"{0:0.0}", (pair.Value as Line).EndPoint.Y) //END.Y
                        , string.Format(@"{0:0.0}", (pair.Value as Line).EndPoint.Z) //END.Z
                        , string.Format(@"{0}", (pair.Value as Line).ColorIndex) //ColorIndex
                        , string.Format(@"{0:0.000}", (pair.Value as Line).Thickness) //Tickness
                    };
                    break;
            }

            return rowRes;
        }

        private static void extractDataWorksheet(ExcelWorksheet ews, GemBox.Spreadsheet.CellRange range, FORMAT format)
        {
            // создать структуру таблицы - добавить поля в таблицу, при необходимости создать таблицу
            createDataTableWorksheet(ews.Name, range, format);

            ews.ExtractDataEvent += new ExcelWorksheet.ExtractDataEventHandler((sender, e) => {
                e.DataTableValue = e.ExcelValue == null ? null : e.ExcelValue.ToString();
                e.Action = ExtractDataEventAction.Continue;
            });

            try {
                ews.ExtractToDataTable(_dictDataTableOfExcelWorksheet[ews.Name]
                    , range.LastRowIndex
                    , ExtractDataOptions.SkipEmptyRows | ExtractDataOptions.StopAtFirstEmptyRow
                    , ews.Rows[range.FirstRowIndex + (format == FORMAT.HEAP ? 0 : format == FORMAT.ORDER ? 1 : 0)]
                    , ews.Columns[range.FirstColumnIndex]
                );
            } catch (Exception e) {
                Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e);

                Logging.AcEditorWriteException(e, string.Format(@"Лист MS Excel: {0}", ews.Name));
            }

            Logging.AcEditorWriteMessage(string.Format(@"На листе с имененм = {0} полей = {1}", ews.Name, range.LastColumnIndex + 1));

            //// добавить записи в таблицу
            //for (i = range.FirstRowIndex + 1; !(i > range.LastRowIndex); i++) {
            //    dataRow = null;

            //    for (j = range.FirstColumnIndex; !(j > range.LastColumnIndex); j++) {
            //        cellValue = string.Empty;
            //        iValidateCellRes = validateCell(range, i, j);

            //        if (iValidateCellRes == VALIDATE_CELL_RESULT.NEW_ROW) {
            //            dataRow = _dictDataTableOfExcelWorksheet[ews.Name].Rows.Add();

            //            cellValue = range[i - range.FirstRowIndex, j - range.FirstColumnIndex].Value.ToString();
            //            acDoc.Editor.WriteMessage(string.Format(@"{0}Добавлена строка для элемента = {1}", Environment.NewLine, cellValue));
            //        } else
            //            if (iValidateCellRes == VALIDATE_CELL_RESULT.CONTINUE)
            //            continue;
            //        else
            //                if (iValidateCellRes == VALIDATE_CELL_RESULT.BREAK)
            //            break;
            //        else
            //            // значение для параметра (VALIDATE_CELL_RESULT.VALUE)
            //            cellValue = range[i - range.FirstRowIndex, j - range.FirstColumnIndex].Value.ToString();

            //        if (dataRow == null)
            //            break;
            //        else
            //            dataRow[j] = cellValue;
            //    }
            //}
        }

        private enum VALIDATE_CELL_RESULT : short { BREAK = -2, CONTINUE, NEW_ROW, VALUE }

        private static VALIDATE_CELL_RESULT validateCell(GemBox.Spreadsheet.CellRange range, int iRow, int iColumn)
        {
            VALIDATE_CELL_RESULT iRes = VALIDATE_CELL_RESULT.BREAK;

            ExcelCell cell = range[iRow - range.FirstRowIndex, iColumn - range.FirstColumnIndex];

            if ((!(cell == null))
                && (!(cell.Value == null)))
                // только при наличии значения
                if (cell.Value.Equals(string.Empty) == true)
                // значение пустое
                    if (iColumn == range.FirstColumnIndex)
                    // если это ИМЯ элемента - запись не добавлять
                        ; // оставить как есть "-1"
                    else
                    // если это параметр - перейти к обработке следующего
                        iRes = VALIDATE_CELL_RESULT.CONTINUE;
                else
                // значение НЕ пустое
                    if (iColumn == range.FirstColumnIndex)
                    // если это ИМЯ элемента - добавить запись
                        iRes = 0;
                    else
                    // если это параметр - присвоить значение для параметра
                        iRes = VALIDATE_CELL_RESULT.VALUE;
            else
            // значения нет
                if (iColumn == range.FirstColumnIndex)
                // если это ИМЯ элемента - запись не добавлять
                    ; // оставить как есть "-1"
                else
                // если это параметр - перейти к обработке следующего
                    iRes = VALIDATE_CELL_RESULT.CONTINUE;

            return iRes;
        }

        private static void closeWorkbook(string strFullNameSettings)
        {
            bool bInstanceAppExcel = true;
            Microsoft.Office.Interop.Excel.Application appExcel = null;

            try {
                appExcel = Marshal.GetActiveObject("Excel.Application") as Microsoft.Office.Interop.Excel.Application;
            }
            catch (Exception) { bInstanceAppExcel = false; }

            if (bInstanceAppExcel == false)
                appExcel = new Microsoft.Office.Interop.Excel.Application();
            else
                ;

            foreach (Microsoft.Office.Interop.Excel.Workbook workbook in appExcel.Workbooks)
                if (workbook.FullName.Equals(strFullNameSettings, StringComparison.InvariantCultureIgnoreCase) == true) {
                    workbook.Saved = true;
                    workbook.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlDoNotSaveChanges);

                    break;
                } else
                    ;
        }

        private static void clearWorkbook(string strNameSettingsExcelFile = @"", FORMAT format = FORMAT.ORDER)
        {
            Database dbCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
            DocumentCollection acDocMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            Document acDoc = acDocMgr.GetDocument(dbCurrent);

            string strNameSettings = getFullNameSettingsExcelFile(strNameSettingsExcelFile);
            int i = -1, j = -1;

            try {
                closeWorkbook(strNameSettings);

                ExcelFile ef = new ExcelFile();
                ef.LoadXls(strNameSettings, XlsOptions.None);

                if (ef.Protected == false) {
                    foreach (ExcelWorksheet ews in ef.Worksheets) {
                        GemBox.Spreadsheet.CellRange range = ews.GetUsedCellRange();

                        acDoc.Editor.WriteMessage(string.Format(@"{0}Очистка листа с имененм = {1}", Environment.NewLine, ews.Name));

                        if (range.LastRowIndex > 0) {
                            // создать структуру таблицы - добавить поля в таблицу, при необходимости создать таблицу
                            createDataTableWorksheet(ews.Name, range);
                            // удалить значения, если есть
                            if (_dictDataTableOfExcelWorksheet[ews.Name].Rows.Count > 0)
                                _dictDataTableOfExcelWorksheet[ews.Name].Rows.Clear();
                            else
                                ;

                            for (i = range.FirstRowIndex + (format == FORMAT.HEAP ? 0 : format == FORMAT.ORDER ? 1 : 0); !(i > range.LastRowIndex); i++) {
                                for (j = range.FirstColumnIndex; !(j > range.LastColumnIndex); j++) {
                                    range[i - range.FirstRowIndex, j - range.FirstColumnIndex].Value = string.Empty;
                                }
                            }
                        } else
                            ;
                    }

                    ef.SaveXls(strNameSettings);
                } else
                    acDoc.Editor.WriteMessage(string.Format(@"{0}Очистка книги с имененм = {1} невозможна", Environment.NewLine, strNameSettings));
            } catch (Exception e) {
                acDoc.Editor.WriteMessage(string.Format(@"{0}Сохранение MSExcel-книги исключение: {1}{0}{2}", Environment.NewLine, e.Message, e.StackTrace));
            }
        }

        private static void saveWorkbook(string strNameSettingsExcelFile = @"", FORMAT format = FORMAT.ORDER)
        {
            Database dbCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
            DocumentCollection acDocMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            Document acDoc = acDocMgr.GetDocument(dbCurrent);

            string strNameSettings = getFullNameSettingsExcelFile(strNameSettingsExcelFile);
            int i = -1, j = -1;

            try {
                closeWorkbook(strNameSettings);

                ExcelFile ef = new ExcelFile();
                ef.LoadXls(strNameSettings, XlsOptions.None);

                if (ef.Protected == false) {
                    foreach (ExcelWorksheet ews in ef.Worksheets) {
                        GemBox.Spreadsheet.CellRange range = ews.GetUsedCellRange();

                        acDoc.Editor.WriteMessage(string.Format(@"{0}Сохранение листа с имененм = {1}", Environment.NewLine, ews.Name));

                        try {
                            ews.InsertDataTable(_dictDataTableOfExcelWorksheet[ews.Name]
                                , range.FirstRowIndex + (format == FORMAT.HEAP ? 0 : format == FORMAT.ORDER ? 1 : 0)
                                , range.FirstColumnIndex
                                , false
                            );

                            //if (range.LastRowIndex > 0) {
                            //    for (i = range.FirstRowIndex + 1; !(i > range.LastRowIndex); i++) {
                            //        for (j = range.FirstColumnIndex; !(j > range.LastColumnIndex); j++) {
                            //            range[i - range.FirstRowIndex, j - range.FirstColumnIndex].Value = string.Empty;
                            //        }
                            //    }
                            //} else
                            //    ;
                        } catch (Exception e) {
                            acDoc.Editor.WriteMessage(string.Format(@"{0}Сохранение книги с имененм = {1}, лист = {2} невозможна", Environment.NewLine, strNameSettings, ews.Name));
                        }
                    }

                    ef.SaveXls(strNameSettings);
                } else
                    acDoc.Editor.WriteMessage(string.Format(@"{0}Сохранение книги с имененм = {1} невозможна", Environment.NewLine, strNameSettings));
            } catch (Exception e) {
                acDoc.Editor.WriteMessage(string.Format(@"{0}Сохранение MSExcel-книги исключение: {1}{0}{2}", Environment.NewLine, e.Message, e.StackTrace));
            }
        }

        private static List<KEY_ENTITY> _listKeyEntityToExport = null;

        public static void AddToExport(Entity entity)
        {
            if (_listKeyEntityToExport == null)
                _listKeyEntityToExport = new List<KEY_ENTITY>();
            else
                ;

            if (s_dictEntity.Values.Contains(entity) == false)
                s_dictEntity.Add(GetKeyEntity(entity.GetType()), entity);
            else
                _listKeyEntityToExport.Add(s_dictEntity.FirstOrDefault(x => x.Value.ObjectId == entity.ObjectId).Key);
        }
        /// <summary>
        /// Экспортировать список объектов в книгу MS Excel
        /// </summary>
        public static int Export(string strNameSettingsExcelFile = @"", FORMAT format = FORMAT.ORDER)
        {
            int iErr = -1;

            object[] dataRow = null;
            string nameWorksheet = string.Empty;
            COMMAND_ENTITY commandEntity = COMMAND_ENTITY.UNKNOWN;
            List<KEY_ENTITY> _listKeyEntityForDelete = null;

            if (!(_listKeyEntityToExport == null)) {
            // удалить лишние элементы
                if (!(s_dictEntity.Keys.Count == _listKeyEntityToExport.Count))
                    if (s_dictEntity.Keys.Count < _listKeyEntityToExport.Count)
                        foreach (KEY_ENTITY key in _listKeyEntityToExport)
                            if (s_dictEntity.Keys.Contains(key) == false)
                                s_dictEntity.Remove(key);
                            else
                                ;
                    else
                        if (s_dictEntity.Keys.Count > _listKeyEntityToExport.Count) {
                            _listKeyEntityForDelete = new List<KEY_ENTITY>();

                            foreach (KEY_ENTITY key in s_dictEntity.Keys)
                                if (_listKeyEntityToExport.IndexOf(key) < 0)
                                    _listKeyEntityForDelete.Add(key);
                                else
                                    ;
                            //???
                            foreach (KEY_ENTITY key in _listKeyEntityForDelete)
                                s_dictEntity.Remove(key);
                        } else
                            ; // других вариантов быть не может
                else
                    ; //??? кол-во равно, но объекты м.б. различные

                _listKeyEntityToExport.Clear();
                _listKeyEntityToExport = null;
            } else
                ;

            try {
                clearWorkbook(strNameSettingsExcelFile, format);

                switch (format) {
                    case FORMAT.HEAP:
                        ExcelFile ef = new ExcelFile();
                        ef.LoadXls(getFullNameSettingsExcelFile(strNameSettingsExcelFile), XlsOptions.None);
                        nameWorksheet = ef.Worksheets[0].Name;
                        break;
                    case FORMAT.ORDER:
                    default:
                        // зависит от типа примитива (будет определена при каждой итерации)
                        break;
                }

                foreach (KeyValuePair<KEY_ENTITY, Entity> pair in s_dictEntity) {
                    switch (format) {
                        case FORMAT.HEAP:
                            // 'nameWorksheet' определено ранее (не зависит от типа примитива)
                            commandEntity = pair.Key.m_command;
                            break;
                        case FORMAT.ORDER:
                        default:
                            nameWorksheet = s_MappingKeyEntity.Find(type => type.m_type == pair.Value.GetType()).Name;
                            commandEntity = (COMMAND_ENTITY)Enum.Parse(typeof(COMMAND_ENTITY), nameWorksheet);
                            break;
                    }

                    if (dictDelegateMethodeEntity.ContainsKey(commandEntity) == true)
                        dataRow = dictDelegateMethodeEntity[commandEntity].entityToDataRow(pair, format);
                    else
                        ;

                    if (!(dataRow == null))
                        _dictDataTableOfExcelWorksheet[nameWorksheet].Rows.Add(dataRow);
                    else
                        ;
                }

                saveWorkbook(strNameSettingsExcelFile, format);

                iErr = 0; // нет ошибок
            } catch (Exception e) {
                iErr = -1;

                Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e);

                Logging.AcEditorWriteException(e, @"Очистка-Преобразование-Сохранение...");
            }

            return iErr;
            }
    }
}
