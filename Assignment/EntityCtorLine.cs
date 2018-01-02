using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Autodesk.Cad.Crushner.Assignment
{
    partial class EntityCtor
    {
        /// <summary>
        /// Создать новый примитив - линия по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - линия</returns>
        public static EntityCtor.ProxyEntity newLineDecart(Settings.EntityParser.ProxyEntity entity/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;
            // соэдать примитив 
            pEntityRes = new ProxyEntity (new Line());
            // значения для параметров примитива

            (pEntityRes.m_entity as Line).StartPoint = new Point3d(
                (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.LINE_START_X)
                , (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.LINE_START_DECART_Y)
                , (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.LINE_START_DECART_Z));
            (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.LINE_END_DECART_X)
                , (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.LINE_END_DECART_Y)
                , (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.LINE_END_DECART_Z));
            (pEntityRes.m_entity as Line).ColorIndex = (int)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.LINE_COLORINDEX);
            (pEntityRes.m_entity as Line).Thickness = (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.LINE_TICKNESS);

            //pEntityRes.m_BlockName = blockName;

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - линия по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - линия</returns>
        public static EntityCtor.ProxyEntity newALineDecartX(Settings.EntityParser.ProxyEntity entity/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;
            // соэдать примитив 
            pEntityRes = new ProxyEntity(new Line());
            // значения для параметров примитива

            (pEntityRes.m_entity as Line).StartPoint = new Point3d(
                (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_START_DECART_X)
                , (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_START_DECART_Y)
                , (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_START_DECART_Z));
            (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                (pEntityRes.m_entity as Line).StartPoint.X +
                    (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_LENGTH)
                , (pEntityRes.m_entity as Line).StartPoint.Y
                , (pEntityRes.m_entity as Line).StartPoint.Z);
            (pEntityRes.m_entity as Line).ColorIndex = int.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_COLORINDEX).ToString());
            (pEntityRes.m_entity as Line).Thickness = (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_TICKNESS);

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - линия по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - линия</returns>
        public static EntityCtor.ProxyEntity newALineDecartY(Settings.EntityParser.ProxyEntity entity/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;
            // соэдать примитив 
            pEntityRes = new ProxyEntity(new Line());
            // значения для параметров примитива
            (pEntityRes.m_entity as Line).StartPoint = new Point3d(
                (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_START_DECART_X)
                , (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_START_DECART_Y)
                , (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_START_DECART_Z));
            (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                (pEntityRes.m_entity as Line).StartPoint.X
                , (pEntityRes.m_entity as Line).StartPoint.Y +
                    (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_LENGTH)
                , (pEntityRes.m_entity as Line).StartPoint.Z);
            (pEntityRes.m_entity as Line).ColorIndex = int.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_COLORINDEX).ToString());
            (pEntityRes.m_entity as Line).Thickness = (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_TICKNESS);

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - линия по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - линия</returns>
        public static EntityCtor.ProxyEntity newALineDecartZ(Settings.EntityParser.ProxyEntity entity/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;
            // соэдать примитив 
            pEntityRes = new ProxyEntity(new Line());
            // значения для параметров примитива
            (pEntityRes.m_entity as Line).StartPoint = new Point3d(
                (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_START_DECART_X)
                , (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_START_DECART_Y)
                , (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_START_DECART_Z));
            (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                (pEntityRes.m_entity as Line).StartPoint.X
                , (pEntityRes.m_entity as Line).StartPoint.Y
                , (pEntityRes.m_entity as Line).StartPoint.Z +
                    (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_LENGTH));
            (pEntityRes.m_entity as Line).ColorIndex = int.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_COLORINDEX).ToString());
            (pEntityRes.m_entity as Line).Thickness = (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ALINE_TICKNESS);

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - линия по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - линия</returns>
        public static EntityCtor.ProxyEntity newRLineDecartX(Settings.EntityParser.ProxyEntity entity/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;
            string nameEntityRelative = string.Empty;
            // соэдать примитив 
            pEntityRes = new ProxyEntity(new Line());
            // значения для параметров примитива
            nameEntityRelative = entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.RLINE_NAME_ENTITY_RELATIVE).ToString();

            (pEntityRes.m_entity as Line).StartPoint = Point3d.Origin;
            (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.RLINE_LENGTH)
                , Point3d.Origin.Y
                , Point3d.Origin.Z
            );

            return pEntityRes;
        }

        public static EntityCtor.ProxyEntity newRLineDecartY(Settings.EntityParser.ProxyEntity entity/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;
            string nameEntityRelative = string.Empty;
            // соэдать примитив 
            pEntityRes = new ProxyEntity(new Line());
            // значения для параметров примитива
            nameEntityRelative = entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.RLINE_NAME_ENTITY_RELATIVE).ToString();

            (pEntityRes.m_entity as Line).StartPoint = Point3d.Origin;
            (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                Point3d.Origin.X
                , (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.RLINE_LENGTH)
                , Point3d.Origin.Z
            );

            return pEntityRes;
        }

        public static EntityCtor.ProxyEntity newRLineDecartZ(Settings.EntityParser.ProxyEntity entity/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;
            string nameEntityRelative = string.Empty;
            // соэдать примитив 
            pEntityRes = new ProxyEntity(new Line());
            // значения для параметров примитива
            nameEntityRelative = entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.RLINE_NAME_ENTITY_RELATIVE).ToString();

            (pEntityRes.m_entity as Line).StartPoint = Point3d.Origin;
            (pEntityRes.m_entity as Line).EndPoint = new Point3d(
                Point3d.Origin.X
                , Point3d.Origin.Y
                , (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.RLINE_LENGTH)
            );

            return pEntityRes;
        }

        public static Settings.EntityParser.ProxyEntity lineDecartToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes;

            rowRes = new Settings.EntityParser.ProxyEntity(new object[] {
                string.Format(@"{0}", pair.Key.Name) //NAME
                , string.Format(@"{0}", pair.Key.Command.ToString()) //!!! COMMAND_ENTITY
                , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).StartPoint.X) //SATRT.X
                , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).StartPoint.Y) //START.Y
                , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).StartPoint.Z) //START.Z
                , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).EndPoint.X) //END.X
                , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).EndPoint.Y) //END.Y
                , string.Format(@"{0:0.0}", (pair.Value.m_entity as Line).EndPoint.Z) //END.Z
                , string.Format(@"{0}", (pair.Value.m_entity as Line).ColorIndex) //ColorIndex
                , string.Format(@"{0:0.000}", (pair.Value.m_entity as Line).Thickness) //Tickness
            });

            return rowRes;
        }

        public static Settings.EntityParser.ProxyEntity lineSphereToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes;

            throw new NotImplementedException();

            return rowRes;
        }
        /// <summary>
        /// Преобразовать объект Acad в объект конфигурации для возможности экспорта
        /// </summary>
        /// <param name="pair">Ключ объекта + объект ACad</param>
        /// <returns>Объект конфигурации</returns>
        public static Settings.EntityParser.ProxyEntity alineDecartXToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes = new Settings.EntityParser.ProxyEntity();

            return rowRes;
        }
        /// <summary>
        /// Преобразовать объект Acad в объект конфигурации для возможности экспорта
        /// </summary>
        /// <param name="pair">Ключ объекта + объект ACad</param>
        /// <returns>Объект конфигурации</returns>
        public static Settings.EntityParser.ProxyEntity alineDecartYToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes = new Settings.EntityParser.ProxyEntity();

            return rowRes;
        }
        /// <summary>
        /// Преобразовать объект Acad в объект конфигурации для возможности экспорта
        /// </summary>
        /// <param name="pair">Ключ объекта + объект ACad</param>
        /// <returns>Объект конфигурации</returns>
        public static Settings.EntityParser.ProxyEntity alineDecartZToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes = new Settings.EntityParser.ProxyEntity();

            return rowRes;
        }
        /// <summary>
        /// Преобразовать объект Acad в объект конфигурации для возможности экспорта
        /// </summary>
        /// <param name="pair">Ключ объекта + объект ACad</param>
        /// <returns>Объект конфигурации</returns>
        public static Settings.EntityParser.ProxyEntity rlineDecartXToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes = new Settings.EntityParser.ProxyEntity();

            return rowRes;
        }
        /// <summary>
        /// Преобразовать объект Acad в объект конфигурации для возможности экспорта
        /// </summary>
        /// <param name="pair">Ключ объекта + объект ACad</param>
        /// <returns>Объект конфигурации</returns>
        public static Settings.EntityParser.ProxyEntity rlineDecartYToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes = new Settings.EntityParser.ProxyEntity();

            return rowRes;
        }
        /// <summary>
        /// Преобразовать объект Acad в объект конфигурации для возможности экспорта
        /// </summary>
        /// <param name="pair">Ключ объекта + объект ACad</param>
        /// <returns>Объект конфигурации</returns>
        public static Settings.EntityParser.ProxyEntity rlineDecartZToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes = new Settings.EntityParser.ProxyEntity();

            return rowRes;
        }
    }
}
