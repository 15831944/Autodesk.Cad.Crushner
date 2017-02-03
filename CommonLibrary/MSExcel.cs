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
using Autodesk.AutoCAD.Interop.Common;

namespace Autodesk.Cad.Crushner.Common
{
    public partial class MSExcel : Collection
    {
        /// <summary>
        /// Известные форматы книги MS Excel
        /// </summary>
        public enum FORMAT : short { UNKNOWN = -1, ORDER, HEAP }
        /// <summary>
        /// Наименование книги MS Excel со списком объектов по умолчанию
        /// </summary>
        private static string s_nameSettings = @"settings.xls";
        /// <summary>
        /// Возвратить полное путь с именем файла (книги MS Excel) конфигурации
        /// </summary>
        /// <param name="strNameSettingsExcelFile"></param>
        /// <returns></returns>
        private static string getFullNameSettingsExcelFile(string strNameSettingsExcelFile = @"")
        {
            return string.Format(@"{0}\{1}"
                //, AppDomain.CurrentDomain.BaseDirectory
                , Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                , strNameSettingsExcelFile.Equals(string.Empty) == false ?
                    strNameSettingsExcelFile :
                        s_nameSettings);
        }
        /// <summary>
        /// Словарь имя_листа::таблица - содержимое книги MS Excel
        /// </summary>
        private static Dictionary<string, System.Data.DataTable> _dictDataTableOfExcelWorksheet;
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
        private delegate EntityParser.ProxyEntity delegateNewEntity(DataRow rEntity, FORMAT format/*, string blockName*/);
        /// <summary>
        /// Тип делегата для упаковки примитива в строку таблицы
        ///  (подготовка к экспорту)
        /// </summary>
        /// <param name="pair">Сложный ключ - идентификатор примитива + объект примитива</param>
        /// <param name="format">Формат книги MS Excel</param>
        /// <returns></returns>
        private delegate object[] delegateEntityToDataRow(KeyValuePair<KEY_ENTITY, EntityParser.ProxyEntity> pair, FORMAT format);
        /// <summary>
        /// Словарь с методами для создания/упаковки прмитивов разных типов
        ///  ключ - тип примитива
        /// </summary>
        private static Dictionary<COMMAND_ENTITY, METHODE_ENTITY> dictDelegateMethodeEntity = new Dictionary<COMMAND_ENTITY, METHODE_ENTITY>() {
            { COMMAND_ENTITY.CIRCLE, new METHODE_ENTITY () { newEntity = EntityParser.newCircle, entityToDataRow = EntityParser.circleToDataRow } }
            , { COMMAND_ENTITY.ARC, new METHODE_ENTITY () { newEntity = EntityParser.newArc, entityToDataRow = EntityParser.arcToDataRow } }
            , { COMMAND_ENTITY.LINE, new METHODE_ENTITY () { newEntity = EntityParser.newLine, entityToDataRow = EntityParser.lineToDataRow } }
            , { COMMAND_ENTITY.PLINE3, new METHODE_ENTITY () { newEntity = EntityParser.newPolyLine3d, entityToDataRow = EntityParser.polyLine3dToDataRow } }
            , { COMMAND_ENTITY.CONE, new METHODE_ENTITY () { newEntity = EntityParser.newCone, entityToDataRow = EntityParser.coneToDataRow } }
            , { COMMAND_ENTITY.BOX, new METHODE_ENTITY () { newEntity = EntityParser.newBox, entityToDataRow = EntityParser.boxToDataRow } }
        };
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

            //GemBox.Spreadsheet.SpreadsheetInfo.SetLicense(@"FREE-LIMITED-KEY");

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

        private static string WSHHEET_NAME_CONFIG = @"_CONFIG";

        private static string WSHHEET_NAME_BLOCK_REFERENCE = @"_BLOCK_REFERENCE";

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
            COMMAND_ENTITY commandEntity = COMMAND_ENTITY.UNKNOWN;
            string nameEntity = string.Empty;

            foreach (ExcelWorksheet ews in ef.Worksheets) {
                if (ews.Name.Equals(WSHHEET_NAME_CONFIG) == false) {
                    range =
                        ews.GetUsedCellRange()
                        //getUsedCellRange(ews, format)
                        ;

                    Logging.AcEditorWriteMessage(string.Format(@"Обработка листа с имененм = {0}", ews.Name));

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
                                    if (dictDelegateTryParseCommandAndNameEntity[format](ews, rEntity, out commandEntity, out nameEntity) == true) {
                                        pEntity = null;

                                        // соэдать примитив 
                                        if (dictDelegateMethodeEntity.ContainsKey(commandEntity) == true)
                                            pEntity = dictDelegateMethodeEntity[commandEntity].newEntity(rEntity, format/*, ews.Name*/);
                                        else
                                            ;

                                        if (!(pEntity == null))
                                            s_dictBlock.AddEntity(
                                                ews.Name
                                                , nameEntity
                                                , pEntity.GetValueOrDefault()
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
                                    , range.LastRowIndex + 1
                                    , _dictDataTableOfExcelWorksheet[ews.Name].Rows.Count
                                ));
                                break;
                        }
                    
                    } else
                    // нет строк с данными
                        Logging.AcEditorWriteMessage(string.Format(@"На листе с имененм = {0} нет строк для распознования", ews.Name));
                } else
                    ; // страница(лист) с конфигурацией 
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

        internal static void AddToExport(object entity)
        {
            throw new NotImplementedException();
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
                    acDoc.Editor.WriteMessage(string.Format(@"{0}Очистка книги с имененм = {1} невозможна", Environment.NewLine, strNameSettings));
            } catch (Exception e) {
                acDoc.Editor.WriteMessage(string.Format(@"{0}Сохранение MSExcel-книги исключение: {1}{0}{2}", Environment.NewLine, e.Message, e.StackTrace));
            }
        }
        /// <summary>
        /// Сохранить внесенные изменения в книге MS Excel
        /// </summary>
        /// <param name="strNameSettingsExcelFile">Полный путь + наименование для сохраняемой книги (файла конфигурации)</param>
        /// <param name="format">Формат книги MS Excel</param>
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

                            //if ((range.LastRowIndex + 1) > 0) {
                            //    for (i = range.FirstRowIndex + 1; !(i > (range.LastRowIndex + 1)); i++) {
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
