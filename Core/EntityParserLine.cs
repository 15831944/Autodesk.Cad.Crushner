using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Autodesk.Cad.Crushner.Core
{
    partial class EntityParser
    {
        /// <summary>
        /// Создать новый примитив - линия по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - линия</returns>
        public static EntityParser.ProxyEntity newLine(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityParser.ProxyEntity pEntityRes;
            // соэдать примитив 
            pEntityRes = new ProxyEntity (new Line());
            // значения для параметров примитива
            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    (pEntityRes.m_entity as Line).StartPoint = new Point3d(
                        double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.LINE_START_X].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.LINE_START_Y].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.LINE_START_Z].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                        double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.LINE_END_X].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.LINE_END_Y].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.LINE_END_Z].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (pEntityRes.m_entity as Line).ColorIndex = int.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.LINE_COLORINDEX].ToString());
                    (pEntityRes.m_entity as Line).Thickness = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.LINE_TICKNESS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);

                    //pEntityRes.m_BlockName = blockName;
                    break;
                case MSExcel.FORMAT.ORDER:
                    (pEntityRes.m_entity as Line).StartPoint = new Point3d(
                        double.Parse(rEntity[@"START.X"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"START.Y"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"START.Z"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                        double.Parse(rEntity[@"END.X"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"END.Y"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"END.Z"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (pEntityRes.m_entity as Line).ColorIndex = int.Parse(rEntity[@"ColorIndex"].ToString());
                    (pEntityRes.m_entity as Line).Thickness = double.Parse(rEntity[@"TICKNESS"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                default:
                    break;
            }

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - линия по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - линия</returns>
        public static EntityParser.ProxyEntity newALineX(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityParser.ProxyEntity pEntityRes;
            // соэдать примитив 
            pEntityRes = new ProxyEntity(new Line());
            // значения для параметров примитива
            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    (pEntityRes.m_entity as Line).StartPoint = new Point3d(
                        double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_START_X].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_START_Y].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_START_Z].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                        (pEntityRes.m_entity as Line).StartPoint.X +
                            double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_LENGTH].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , (pEntityRes.m_entity as Line).StartPoint.Y
                        , (pEntityRes.m_entity as Line).StartPoint.Z);
                    (pEntityRes.m_entity as Line).ColorIndex = int.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_COLORINDEX].ToString());
                    (pEntityRes.m_entity as Line).Thickness = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_TICKNESS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                default:
                    break;
            }

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - линия по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - линия</returns>
        public static EntityParser.ProxyEntity newALineY(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityParser.ProxyEntity pEntityRes;
            // соэдать примитив 
            pEntityRes = new ProxyEntity(new Line());
            // значения для параметров примитива
            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    (pEntityRes.m_entity as Line).StartPoint = new Point3d(
                        double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_START_X].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_START_Y].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_START_Z].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                        (pEntityRes.m_entity as Line).StartPoint.X
                        , (pEntityRes.m_entity as Line).StartPoint.Y +
                            double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_LENGTH].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , (pEntityRes.m_entity as Line).StartPoint.Z);
                    (pEntityRes.m_entity as Line).ColorIndex = int.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_COLORINDEX].ToString());
                    (pEntityRes.m_entity as Line).Thickness = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_TICKNESS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                default:
                    break;
            }

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - линия по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - линия</returns>
        public static EntityParser.ProxyEntity newALineZ(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityParser.ProxyEntity pEntityRes;
            // соэдать примитив 
            pEntityRes = new ProxyEntity(new Line());
            // значения для параметров примитива
            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    (pEntityRes.m_entity as Line).StartPoint = new Point3d(
                        double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_START_X].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_START_Y].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_START_Z].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                        (pEntityRes.m_entity as Line).StartPoint.X
                        , (pEntityRes.m_entity as Line).StartPoint.Y
                        , (pEntityRes.m_entity as Line).StartPoint.Z +
                            double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_LENGTH].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (pEntityRes.m_entity as Line).ColorIndex = int.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_COLORINDEX].ToString());
                    (pEntityRes.m_entity as Line).Thickness = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ALINE_TICKNESS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                default:
                    break;
            }

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - линия по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - линия</returns>
        public static EntityParser.ProxyEntity newRLineX(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityParser.ProxyEntity pEntityRes;
            string nameEntityRelative = string.Empty;
            // соэдать примитив 
            pEntityRes = new ProxyEntity(new Line());
            // значения для параметров примитива
            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    nameEntityRelative = rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.RLINE_NAME_ENTITY_RELATIVE].ToString();

                    (pEntityRes.m_entity as Line).StartPoint = Point3d.Origin;
                    (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                        double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.RLINE_LENGTH].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , Point3d.Origin.Y
                        , Point3d.Origin.Z
                    );
                    break;
                default:
                    break;
            }

            return pEntityRes;
        }

        public static EntityParser.ProxyEntity newRLineY(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityParser.ProxyEntity pEntityRes;
            string nameEntityRelative = string.Empty;
            // соэдать примитив 
            pEntityRes = new ProxyEntity(new Line());
            // значения для параметров примитива
            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    nameEntityRelative = rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.RLINE_NAME_ENTITY_RELATIVE].ToString();

                    (pEntityRes.m_entity as Line).StartPoint = Point3d.Origin;
                    (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                        Point3d.Origin.X
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.RLINE_LENGTH].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , Point3d.Origin.Z
                    );
                    break;
                default:
                    break;
            }

            return pEntityRes;
        }

        public static EntityParser.ProxyEntity newRLineZ(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityParser.ProxyEntity pEntityRes;
            string nameEntityRelative = string.Empty;
            // соэдать примитив 
            pEntityRes = new ProxyEntity(new Line());
            // значения для параметров примитива
            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    nameEntityRelative = rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.RLINE_NAME_ENTITY_RELATIVE].ToString();

                    (pEntityRes.m_entity as Line).StartPoint = Point3d.Origin;
                    (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                        Point3d.Origin.X
                        , Point3d.Origin.Y
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.RLINE_LENGTH].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)                        
                    );
                    break;
                default:
                    break;
            }

            return pEntityRes;
        }

        public static object[] lineToDataRow(KeyValuePair<KEY_ENTITY, EntityParser.ProxyEntity> pair, MSExcel.FORMAT format)
        {
            object[] rowRes = null;

            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0}", pair.Key.m_command.ToString()) //!!! COMMAND_ENTITY
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).StartPoint.X) //SATRT.X
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).StartPoint.Y) //START.Y
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).StartPoint.Z) //START.Z
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).EndPoint.X) //END.X
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).EndPoint.Y) //END.Y
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).EndPoint.Z) //END.Z
                        , string.Format(@"{0}", (pair.Value.m_entity as Line).ColorIndex) //ColorIndex
                        , string.Format(@"{0:0.000}", (pair.Value.m_entity as Line).Thickness) //Tickness
                    };
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).StartPoint.X) //SATRT.X
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).StartPoint.Y) //START.Y
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).StartPoint.Z) //START.Z
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).EndPoint.X) //END.X
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).EndPoint.Y) //END.Y
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).EndPoint.Z) //END.Z
                        , string.Format(@"{0}", (pair.Value.m_entity as Line).ColorIndex) //ColorIndex
                        , string.Format(@"{0:0.000}", (pair.Value.m_entity as Line).Thickness) //Tickness
                    };
                    break;
            }

            return rowRes;
        }

        public static object[] alineXToDataRow(KeyValuePair<KEY_ENTITY, EntityParser.ProxyEntity> pair, MSExcel.FORMAT format)
        {
            object[] rowRes = null;

            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    break;
            }

            return rowRes;
        }

        public static object[] alineYToDataRow(KeyValuePair<KEY_ENTITY, EntityParser.ProxyEntity> pair, MSExcel.FORMAT format)
        {
            object[] rowRes = null;

            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    break;
            }

            return rowRes;
        }

        public static object[] alineZToDataRow(KeyValuePair<KEY_ENTITY, EntityParser.ProxyEntity> pair, MSExcel.FORMAT format)
        {
            object[] rowRes = null;

            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    break;
            }

            return rowRes;
        }

        public static object[] rlineXToDataRow(KeyValuePair<KEY_ENTITY, EntityParser.ProxyEntity> pair, MSExcel.FORMAT format)
        {
            object[] rowRes = null;

            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    break;
            }

            return rowRes;
        }

        public static object[] rlineYToDataRow(KeyValuePair<KEY_ENTITY, EntityParser.ProxyEntity> pair, MSExcel.FORMAT format)
        {
            object[] rowRes = null;

            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    break;
            }

            return rowRes;
        }

        public static object[] rlineZToDataRow(KeyValuePair<KEY_ENTITY, EntityParser.ProxyEntity> pair, MSExcel.FORMAT format)
        {
            object[] rowRes = null;

            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    break;
            }

            return rowRes;
        }
    }
}
