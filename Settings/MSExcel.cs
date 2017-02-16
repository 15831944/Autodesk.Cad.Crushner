using Autodesk.Cad.Crushner.Core;
using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace Autodesk.Cad.Crushner.Settings
{
    public partial class MSExcel
    {
        /// <summary>
        /// Перечисление - типы примитивов(сущностей)
        /// </summary>
        public enum COMMAND_ENTITY : short
        {
            UNKNOWN = 0
            , CIRCLE, ARC, LINE_DECART, PLINE3, CONE, BOX
            , ALINE_DECART_X, ALINE_DECART_Y, ALINE_DECART_Z, RLINE_DECART_X, RLINE_DECART_Y, RLINE_DECART_Z, LINE_SPHERE
                , COUNT
        }
        /// <summary>
        /// Индекс столбца на листе книги MSExcel с которого начинаются значения свойств сущностей для отображения
        /// </summary>
        public const int _HEAP_INDEX_COLUMN_PROPERTY = 2;
        /// <summary>
        /// Перечисление - индексы столбцов на листе книги MS Excel в формате 'HEAP'
        /// </summary>
        public enum HEAP_INDEX_COLUMN
        {
            CIRCLE_CENTER_X = _HEAP_INDEX_COLUMN_PROPERTY, CIRCLE_CENTER_Y, CIRCLE_CENTER_Z, CIRCLE_RADIUS, CIRCLE_COLORINDEX, CIRCLE_TICKNESS
            , ARC_CENTER_X = _HEAP_INDEX_COLUMN_PROPERTY, ARC_CENTER_Y, ARC_CENTER_Z, ARC_RADIUS, ARC_ANGLE_START, ARC_ANGLE_END, ARC_COLORINDEX, ARC_TICKNESS
            , LINE_DECART_START_X = _HEAP_INDEX_COLUMN_PROPERTY, LINE_START_DECART_Y, LINE_START_DECART_Z, LINE_END_DECART_X, LINE_END_DECART_Y, LINE_END_DECART_Z, LINE_DECART_COLORINDEX, LINE_DECART_TICKNESS
            , POLYLINE_X_START = _HEAP_INDEX_COLUMN_PROPERTY
            , CONE_HEIGHT = _HEAP_INDEX_COLUMN_PROPERTY, CONE_ARADIUS_X, CONE_ARADIUS_Y, CONE_RADIUS_TOP, CONE_PTDISPLACEMENT_X, CONE_PTDISPLACEMENT_Y, CONE_PTDISPLACEMENT_Z
            , BOX_LAENGTH_X = _HEAP_INDEX_COLUMN_PROPERTY, BOX_LAENGTH_Y, BOX_LAENGTH_Z, BOX_PTDISPLACEMENT_X, BOX_PTDISPLACEMENT_Y, BOX_PTDISPLACEMENT_Z
            , ALINE_START_DECART_X = _HEAP_INDEX_COLUMN_PROPERTY, ALINE_START_DECART_Y, ALINE_START_DECART_Z, ALINE_DECART_LENGTH, ALINE_DECART_COLORINDEX, ALINE_DECART_TICKNESS
            , RLINE_DECART_NAME_ENTITY_RELATIVE = _HEAP_INDEX_COLUMN_PROPERTY, RLINE_DECART_LENGTH, RLINE_DECART_COLORINDEX, RLINE_DECART_TICKNESS
            , LINE_SPHERE_START_X, LINE_SPHERE_START_Y, LINE_SPHERE_START_Z, LINE_SPHERE_RADIUS, LINE_SPHERE_ANGLE_XY, LINE_SPHERE_ANGLE_Z, LINE_SPHERE_COLORINDEX, LINE_SPHERE_TICKNESS
            ,
        }
        /// <summary>
        /// Известные форматы книги MS Excel
        /// </summary>
        public enum FORMAT : short {
            /// <summary>Неизвестный формат</summary>
            UNKNOWN = -1,
            /// <summary>Формат при котором столбцы поименованы</summary>
            ORDER,
            /// <summary>Формат при котором столбцы не поименованы, следуют один за другим</summary>
            HEAP
        }
        /// <summary>
        /// Наименование книги MS Excel со списком объектов по умолчанию
        /// </summary>
        private static string s_nameSettings = @"settings.xls";

        private delegate EntityParser.ProxyEntity delegateNewEntity(DataRow rEntity, FORMAT format/*, string blockName*/);

        private static Dictionary<COMMAND_ENTITY, delegateNewEntity> dictDelegateNewProxyEntity = new Dictionary<COMMAND_ENTITY, delegateNewEntity>() {
            { COMMAND_ENTITY.ARC, EntityParser.newArc }
            , { COMMAND_ENTITY.CIRCLE, EntityParser.newCircle }
            , { COMMAND_ENTITY.LINE_DECART, EntityParser.newLineDecart }
            , { COMMAND_ENTITY.PLINE3, EntityParser.newPolyLine3d }
            , { COMMAND_ENTITY.BOX, EntityParser.newBox }
            , { COMMAND_ENTITY.CONE, EntityParser.newCone }
            // векторы - линиии
            , { COMMAND_ENTITY.ALINE_DECART_X, EntityParser.newALineX }
            , { COMMAND_ENTITY.ALINE_DECART_Y, EntityParser.newALineY }
            , { COMMAND_ENTITY.ALINE_DECART_Z, EntityParser.newALineZ }
            , { COMMAND_ENTITY.RLINE_DECART_X, EntityParser.newRLineX }
            , { COMMAND_ENTITY.RLINE_DECART_Y, EntityParser.newRLineY }
            , { COMMAND_ENTITY.RLINE_DECART_Z, EntityParser.newRLineZ }
            // линия - сферическая СК
            , { COMMAND_ENTITY.LINE_SPHERE, EntityParser.newLineSphere }
            ,
        };
        /// <summary>
        /// Возвратить полное путь с именем файла (книги MS Excel) конфигурации
        /// </summary>
        /// <param name="strNameSettingsExcelFile"></param>
        /// <returns></returns>
        private static string getFullNameSettingsExcelFile(string strNameSettingsExcelFile = @"")
        {
            //var assembly = Assembly.GetAssembly(typeof(MSExcel));
            //var assemblyFileUri = new Uri(assembly.CodeBase);
            //var path = assemblyFileUri.LocalPath;

            return string.Format(@"{0}\{1}"
                //, AppDomain.CurrentDomain.BaseDirectory
                , Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().Location
                    //assemblyFileUri.LocalPath
                )
                , strNameSettingsExcelFile.Equals(string.Empty) == false ?
                    strNameSettingsExcelFile :
                        s_nameSettings);
        }
        /// <summary>
        /// Словарь имя_листа::таблица - содержимое книги MS Excel
        /// </summary>
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
        /// Словарь с определениями блоков
        /// </summary>
        public static MSExcel.DictionaryBlock s_dictBlock = new MSExcel.DictionaryBlock();

        private static string WSHHEET_NAME_CONFIG = @"_CONFIG";

        private static string WSHHEET_NAME_BLOCK_REFERENCE = @"_BLOCK_REFERENCE";

        public static object COORD3d { get; private set; }

        /// <summary>
        /// Импортировать список объектов
        /// </summary>
        /// <param name="strNameSettingsExcelFile">Наименование файла конфигурации (книги MS Excel)</param>>
        /// /// <param name="format">Формат файла конфигурации (книги MS Excel)</param>>
        /// <returns>Признак ошибки при выполнении метода</returns>
        public static int Import(string strNameSettingsExcelFile = @"", Settings.MSExcel.FORMAT format = Settings.MSExcel.FORMAT.ORDER)
        {
            int iErr = 0;

            string strNameSettings = getFullNameSettingsExcelFile(strNameSettingsExcelFile);

            _dictDataTableOfExcelWorksheet = new Dictionary<string, System.Data.DataTable>();

            //GemBox.Spreadsheet.SpreadsheetInfo.SetLicense(@"FREE-LIMITED-KEY");

            try
            {
                ExcelFile ef = new ExcelFile();
                ef.LoadXls(strNameSettings, XlsOptions.None);

                Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Книга открыта {0}, листов = {1}", strNameSettings, ef.Worksheets.Count));

                iErr = import(ef, format);
            } catch (Exception e) {
                iErr = -1;

                Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e);
            }

            return iErr;
        }

        private enum INDEX_COORD3d { UNKNOWN = -2, ANY, X, Y, Z }

        public static int GetLineEndPoint3d(COMMAND_ENTITY commandSlave, string blockName, string nameEntity, out POINT3D value)
        {
            int iRes = 0;
            value = new POINT3D();

            EntityParser.ProxyEntity entityFind;
            // ??? проверить на корректность связи ведущий-ведомый (есть ли сопряжение по оси)
            INDEX_COORD3d indxCoord3dLeading = INDEX_COORD3d.UNKNOWN
                , indxCoord3dSlave = INDEX_COORD3d.UNKNOWN;

            switch (commandSlave) {
                //case COMMAND_ENTITY.ALINE_X:
                case COMMAND_ENTITY.RLINE_DECART_X:
                    indxCoord3dSlave = INDEX_COORD3d.X;
                    break;
                //case COMMAND_ENTITY.ALINE_Y:
                case COMMAND_ENTITY.RLINE_DECART_Y:
                    indxCoord3dSlave = INDEX_COORD3d.Z;
                    break;
                //case COMMAND_ENTITY.ALINE_Z:
                case COMMAND_ENTITY.RLINE_DECART_Z:
                    indxCoord3dSlave = INDEX_COORD3d.Z;
                    break;
                default:
                    break;
            }

            if (s_dictBlock.ContainsKey(blockName) == true) {
                entityFind = s_dictBlock[blockName].m_dictEntityParser.Values.First(entity => {
                    return entity.m_name.Equals(nameEntity) == true;
                });

                if (!(entityFind.m_command == COMMAND_ENTITY.UNKNOWN)) {
                    switch (entityFind.m_command) {
                        case COMMAND_ENTITY.LINE_DECART:
                            indxCoord3dLeading = INDEX_COORD3d.ANY;
                            break;
                        case COMMAND_ENTITY.ALINE_DECART_X:
                        //case COMMAND_ENTITY.RLINE_X:
                            indxCoord3dLeading = INDEX_COORD3d.X;
                            break;
                        case COMMAND_ENTITY.ALINE_DECART_Y:
                        //case COMMAND_ENTITY.RLINE_Y:
                            indxCoord3dLeading = INDEX_COORD3d.Z;
                            break;
                        case COMMAND_ENTITY.ALINE_DECART_Z:
                        //case COMMAND_ENTITY.RLINE_Z:
                            indxCoord3dLeading = INDEX_COORD3d.Z;
                            break;
                        default:
                            break;
                    }

                    iRes = ((indxCoord3dLeading == indxCoord3dSlave) || (indxCoord3dLeading == INDEX_COORD3d.ANY)) ? 0 : -1;

                    if (iRes == 0)
                        value = new POINT3D(new double[] {
                            (double)entityFind.GetProperty(MSExcel.HEAP_INDEX_COLUMN.LINE_END_DECART_X)
                            , (double)entityFind.GetProperty(MSExcel.HEAP_INDEX_COLUMN.LINE_END_DECART_Y)
                            , (double)entityFind.GetProperty(MSExcel.HEAP_INDEX_COLUMN.LINE_END_DECART_Z)
                        });
                } else
                    iRes = -2;
            } else
                iRes = -3;

            return iRes;
        }
        /// <summary>
        /// Импортировать список объектов
        /// </summary>
        /// <param name="ef">Объект книги MS Excel</param>
        /// <param name="format">Формат книги MS Excel</param>
        /// <returns>Признак результата выполнения метода</returns>
        private static int import(ExcelFile ef, FORMAT format)
        {
            int iRes = 0;

            GemBox.Spreadsheet.CellRange range;
            EntityParser.ProxyEntity? pEntity;
            string nameEntity = string.Empty;
            COMMAND_ENTITY commandEntity = COMMAND_ENTITY.UNKNOWN;

            foreach (ExcelWorksheet ews in ef.Worksheets) {
                if (ews.Name.Equals(WSHHEET_NAME_CONFIG) == false) {
                    range =
                        ews.GetUsedCellRange()
                        //getUsedCellRange(ews, format)
                        ;

                    Core.Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Обработка листа с имененм = {0}", ews.Name));

                    if ((!(range == null))
                        && ((range.LastRowIndex + 1)  > 0)) {
                        extractDataWorksheet(ews, range, format);

                        switch (ef.Worksheets.Cast<ExcelWorksheet>().ToList().IndexOf(ews)) {
                            case 0: // BLOCK
                                foreach (DataRow rReferenceBlock in _dictDataTableOfExcelWorksheet[ews.Name].Rows)
                                    s_dictBlock.AddReference(rReferenceBlock);
                                break;
                            default:
                                foreach (DataRow rEntity in _dictDataTableOfExcelWorksheet[ews.Name].Rows) {
                                    if (EntityParser.TryParseCommandAndNameEntity(format, rEntity, out nameEntity, out commandEntity) == true) {
                                        pEntity = null;

                                        // соэдать примитив 
                                        if (dictDelegateNewProxyEntity.ContainsKey(commandEntity) == true)
                                            pEntity = dictDelegateNewProxyEntity[commandEntity](rEntity, format/*, ews.Name*/);
                                        else
                                            ;

                                        if (!(pEntity == null))
                                            s_dictBlock.AddEntity(
                                                ews.Name
                                                , nameEntity
                                                , pEntity.GetValueOrDefault()
                                            );
                                        else
                                            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Элемент с именем {0} пропущен..."
                                                , nameEntity
                                            ));
                                    } else
                                    // ошибка при получении типа и наименования примитива
                                        Core.Logging.DebugCaller(MethodBase.GetCurrentMethod()
                                            , string.Format(@"Ошибка опрделения имени, типа  сущности лист={0}, строка={1}...", ews.Name, _dictDataTableOfExcelWorksheet[ews.Name].Rows.IndexOf(rEntity)));
                                } // цикл по строкам таблицы для листа книги MS Excel

                                Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"На листе с имененм = {0} обработано строк = {1}, добавлено элементов {2}"
                                    , ews.Name
                                    , range.LastRowIndex + 1
                                    , _dictDataTableOfExcelWorksheet[ews.Name].Rows.Count
                                ));
                                break;
                        }
                    
                    } else
                    // нет строк с данными
                        Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"На листе с имененм = {0} нет строк для распознования", ews.Name));
                } else
                    ; // страница(лист) с конфигурацией 
            }

            return iRes;
        }

        private static GemBox.Spreadsheet.CellRange getUsedCellRange(ExcelWorksheet ews, FORMAT format)
        {
            GemBox.Spreadsheet.CellRange rangeRes = null;

            int iRow = -1
                , iColumn = -1;

            // только для 'FORMAT.HEAP'
            iRow = 0;
            iColumn = 0;
            while (!(ews.Rows[iRow].Cells[0].Value == null)) {
                while (!(ews.Rows[iRow].Cells[iColumn].Value == null))
                    iColumn++;

                iRow++;
            }

            if ((iRow > 0)
                && (iColumn > 0))
                rangeRes = ews.Cells.GetSubrangeAbsolute(0, 0, iRow - 1, iColumn - 1);
            else
                ;

            return rangeRes;
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
                    , range.LastRowIndex + 1
                    , ExtractDataOptions.SkipEmptyRows | ExtractDataOptions.StopAtFirstEmptyRow
                    , ews.Rows[range.FirstRowIndex + (format == FORMAT.HEAP ? 0 : format == FORMAT.ORDER ? 1 : 0)]
                    , ews.Columns[range.FirstColumnIndex]
                );

                _dictDataTableOfExcelWorksheet[ews.Name].TableName = ews.Name;
            } catch (Exception e) {
                Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e, string.Format(@"Лист MS Excel: {0}", ews.Name));
            }

            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"На листе с имененм = {0} полей = {1}", ews.Name, range.LastColumnIndex + 1));

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
        /// <summary>
        /// Идентификаторы результата проверки значения в ячейке
        /// </summary>
        private enum VALIDATE_CELL_RESULT : short { BREAK = -2, CONTINUE, NEW_ROW, VALUE }
        /// <summary>
        /// Проверить значение в ячейке
        /// </summary>
        /// <param name="range">Диапазон столбцов/строк</param>
        /// <param name="iRow">Номер строки</param>
        /// <param name="iColumn">Номер столбца</param>
        /// <returns>Результат проверки</returns>
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
        /// <summary>
        /// Закрыть книгу MS Excel
        /// </summary>
        /// <param name="strFullName">Полный путь + наименование для закрываемой книги</param>
        private static void closeWorkbook(string strFullName)
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
                if (workbook.FullName.Equals(strFullName, StringComparison.InvariantCultureIgnoreCase) == true) {
                    workbook.Saved = true;
                    workbook.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlDoNotSaveChanges);

                    break;
                } else
                    ;
        }
        /// <summary>
        /// Очитсить содержимое книги MS Excel
        /// </summary>
        /// <param name="strNameSettingsExcelFile">Полный путь + наименование для очищаемой книги (файла конфигурации)</param>
        /// <param name="format">Формат книги MS Excel</param>
        private static void clearWorkbook(string strNameSettingsExcelFile = @"", FORMAT format = FORMAT.ORDER)
        {
            string strNameSettings = getFullNameSettingsExcelFile(strNameSettingsExcelFile);
            int i = -1, j = -1;

            try {
                closeWorkbook(strNameSettings);

                ExcelFile ef = new ExcelFile();
                ef.LoadXls(strNameSettings, XlsOptions.None);

                if (ef.Protected == false) {
                    foreach (ExcelWorksheet ews in ef.Worksheets) {
                        GemBox.Spreadsheet.CellRange range = ews.GetUsedCellRange();

                        Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"{0}Очистка листа с имененм = {1}", Environment.NewLine, ews.Name));

                        if ((range.LastRowIndex + 1) > 0) {
                            // создать структуру таблицы - добавить поля в таблицу, при необходимости создать таблицу
                            createDataTableWorksheet(ews.Name, range);
                            // удалить значения, если есть
                            if (_dictDataTableOfExcelWorksheet[ews.Name].Rows.Count > 0)
                                _dictDataTableOfExcelWorksheet[ews.Name].Rows.Clear();
                            else
                                ;

                            for (i = range.FirstRowIndex + (format == FORMAT.HEAP ? 0 : format == FORMAT.ORDER ? 1 : 0); !(i > (range.LastRowIndex + 1)); i++) {
                                for (j = range.FirstColumnIndex; !(j > range.LastColumnIndex); j++) {
                                    range[i - range.FirstRowIndex, j - range.FirstColumnIndex].Value = string.Empty;
                                }
                            }
                        } else
                            ;
                    }

                    ef.SaveXls(strNameSettings);
                } else
                    Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"{0}Очистка книги с имененм = {1} невозможна", Environment.NewLine, strNameSettings));
            } catch (Exception e) {
                Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"{0}Сохранение MSExcel-книги исключение: {1}{0}{2}", Environment.NewLine, e.Message, e.StackTrace));
            }
        }
        /// <summary>
        /// Сохранить внесенные изменения в книге MS Excel
        /// </summary>
        /// <param name="strNameSettingsExcelFile">Полный путь + наименование для сохраняемой книги (файла конфигурации)</param>
        /// <param name="format">Формат книги MS Excel</param>
        private static void saveWorkbook(string strNameSettingsExcelFile = @"", FORMAT format = FORMAT.ORDER)
        {
            string strNameSettings = getFullNameSettingsExcelFile(strNameSettingsExcelFile);
            int i = -1, j = -1;

            try {
                closeWorkbook(strNameSettings);

                ExcelFile ef = new ExcelFile();
                ef.LoadXls(strNameSettings, XlsOptions.None);

                if (ef.Protected == false) {
                    foreach (ExcelWorksheet ews in ef.Worksheets) {
                        GemBox.Spreadsheet.CellRange range = ews.GetUsedCellRange();

                        Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"{0}Сохранение листа с имененм = {1}", Environment.NewLine, ews.Name));

                        try {
                            ews.InsertDataTable(_dictDataTableOfExcelWorksheet[ews.Name]
                                , range.FirstRowIndex + (format == FORMAT.HEAP ? 0 : format == FORMAT.ORDER ? 1 : 0)
                                , range.FirstColumnIndex
                                , false
                            );

                            //if ((range.LastRowIndex + 1) > 0) {
                            //    for (i = range.FirstRowIndex + 1; !(i > (range.LastRowIndex + 1)); i++) {
                            //        for (j = range.FirstColumnIndex; !(j > range.LastColumnIndex); j++) {
                            //            range[i - range.FirstRowIndex, j - range.FirstColumnIndex].Value = string.Empty;
                            //        }
                            //    }
                            //} else
                            //    ;
                        } catch (Exception e) {
                            Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"{0}Сохранение книги с имененм = {1}, лист = {2} невозможна", Environment.NewLine, strNameSettings, ews.Name));
                        }
                    }

                    ef.SaveXls(strNameSettings);
                } else
                    Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"{0}Сохранение книги с имененм = {1} невозможна", Environment.NewLine, strNameSettings));
            } catch (Exception e) {
                Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"{0}Сохранение MSExcel-книги исключение: {1}{0}{2}", Environment.NewLine, e.Message, e.StackTrace));
            }
        }

        public static void Clear()
        {
            s_dictBlock?.Clear();
        }

        #region ??? Export - не реализован
        ///// <summary>
        ///// Список примитивов, подготовленный для экспорта (что есть на чертеже)
        /////  , для сравнения с импортированным ранее; в случае разницы добавлять/удалять примитивы
        ///// </summary>
        //private static List<KEY_ENTITY> _listKeyEntityToExport = null;
        ///// <summary>
        ///// Добавить примитив в список, подготовленный для экспорта (что есть на чертеже)
        ///// </summary>
        ///// <param name="pEntity">Примитив для добавления</param>
        //public static void AddToExport(EntityParser.ProxyEntity pEntity)
        //{
        //    if (_listKeyEntityToExport == null)
        //        _listKeyEntityToExport = new List<KEY_ENTITY>();
        //    else
        //        ;

        //    if (s_dictBlock.AddToExport(pEntity) == true)
        //        _listKeyEntityToExport.Add(s_dictBlock.GetKeyEntity (pEntity));
        //    else
        //        ;
        //}
        ///// <summary>
        ///// Экспортировать список объектов в книгу MS Excel
        ///// </summary>
        //public static int Export(string strNameSettingsExcelFile = @"", FORMAT format = FORMAT.ORDER)
        //{
        //    int iErr = -1;

        //    object[] dataRow = null;
        //    string nameWorksheet = string.Empty;
        //    COMMAND_ENTITY commandEntity = COMMAND_ENTITY.UNKNOWN;
        //    List<KEY_ENTITY> _listKeyEntityForDelete = null;

        //    if (!(_listKeyEntityToExport == null)) {
        //    // удалить лишние элементы
        //        if (!(s_dictEntity.Keys.Count == _listKeyEntityToExport.Count))
        //            if (s_dictEntity.Keys.Count < _listKeyEntityToExport.Count)
        //                foreach (KEY_ENTITY key in _listKeyEntityToExport)
        //                    if (s_dictEntity.Keys.Contains(key) == false)
        //                        s_dictEntity.Remove(key);
        //                    else
        //                        ;
        //            else
        //                if (s_dictEntity.Keys.Count > _listKeyEntityToExport.Count) {
        //                    _listKeyEntityForDelete = new List<KEY_ENTITY>();

        //                    foreach (KEY_ENTITY key in s_dictEntity.Keys)
        //                        if (_listKeyEntityToExport.IndexOf(key) < 0)
        //                            _listKeyEntityForDelete.Add(key);
        //                        else
        //                            ;
        //                    //???
        //                    foreach (KEY_ENTITY key in _listKeyEntityForDelete)
        //                        s_dictEntity.Remove(key);
        //                } else
        //                    ; // других вариантов быть не может
        //        else
        //            ; //??? кол-во равно, но объекты м.б. различные

        //        _listKeyEntityToExport.Clear();
        //        _listKeyEntityToExport = null;
        //    } else
        //        ;

        //    try {
        //        clearWorkbook(strNameSettingsExcelFile, format);

        //        switch (format) {
        //            case FORMAT.HEAP:
        //                ExcelFile ef = new ExcelFile();
        //                ef.LoadXls(getFullNameSettingsExcelFile(strNameSettingsExcelFile), XlsOptions.None);
        //                nameWorksheet = ef.Worksheets[0].Name;
        //                break;
        //            case FORMAT.ORDER:
        //            default:
        //                // зависит от типа примитива (будет определена при каждой итерации)
        //                break;
        //        }

        //        foreach (KeyValuePair<KEY_ENTITY, EntityParser.ProxyEntity> pair in s_dictEntity) {
        //            switch (format) {
        //                case FORMAT.HEAP:
        //                    // 'nameWorksheet' определено ранее (не зависит от типа примитива)
        //                    commandEntity = pair.Key.m_command;
        //                    break;
        //                case FORMAT.ORDER:
        //                default:
        //                    nameWorksheet = s_MappingKeyEntity.Find(type => type.m_type == pair.Value.GetType()).Name;
        //                    commandEntity = (COMMAND_ENTITY)Enum.Parse(typeof(COMMAND_ENTITY), nameWorksheet);
        //                    break;
        //            }

        //            if (dictDelegateMethodeEntity.ContainsKey(commandEntity) == true)
        //                dataRow = dictDelegateMethodeEntity[commandEntity].entityToDataRow(pair, format);
        //            else
        //                ;

        //            if (!(dataRow == null))
        //                _dictDataTableOfExcelWorksheet[nameWorksheet].Rows.Add(dataRow);
        //            else
        //                ;
        //        }

        //        saveWorkbook(strNameSettingsExcelFile, format);

        //        iErr = 0; // нет ошибок
        //    } catch (Exception e) {
        //        iErr = -1;

        //        Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e);

        //        Logging.AcEditorWriteException(e, @"Очистка-Преобразование-Сохранение...");
        //    }

        //    return iErr;
        //}
        #endregion
    }
}
