using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Autodesk.Cad.Crushner.Settings
{
    /// <summary>
    /// Преобразователь свойства сущности - запись в таблице БД и обратно
    /// </summary>
    public partial class EntityParser
    {
        private enum INDEX_KEY { NAME, COMMAND }

        public static bool TryParseCommandAndNameEntity(MSExcel.FORMAT format, DataRow rEntity, out string name, out MSExcel.COMMAND_ENTITY command)
        {
            bool bRes = false;

            name = string.Empty;
            command = MSExcel.COMMAND_ENTITY.UNKNOWN;

            switch(format) {
                case MSExcel.FORMAT.HEAP:
                    name = (string)rEntity[(int)INDEX_KEY.NAME];
                    bRes = Enum.TryParse((string)rEntity[(int)INDEX_KEY.COMMAND], true, out command);
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    break;
            }

            return bRes;
        }
        /// <summary>
        /// Свойства сущности из записи в таблице БД
        /// </summary>
        public struct ProxyEntity
        {
            /// <summary>
            /// Свойство сущности для отображения
            /// </summary>
            public struct Property
            {
                public int Index;

                public object Value;
                /// <summary>
                /// Конструктор - основной (с параметрами)
                /// </summary>
                /// <param name="indx">Индекс значения в записи таблитцы БД конфигурации</param>
                /// <param name="value">Значение свойства</param>
                public Property(int indx, object value)
                {
                    Index = indx;

                    Value = value;
                }
            }
            /// <summary>
            /// Наименование сущности, уникальное в пределах блока
            /// </summary>
            public string m_name;
            /// <summary>
            /// Команда для создания сущности
            /// </summary>
            public MSExcel.COMMAND_ENTITY m_command;
            /// <summary>
            /// Значения свойств сущности
            /// </summary>
            public Property[] Properties;
            /// <summary>
            /// Возвратить значение свойства по индексу (номеру столбца в файле конфигурации)
            /// </summary>
            /// <param name="indx">Индекс значения (№ столбца в файле конфигурации - книге MS Excel)</param>
            /// <returns>Значение свойства</returns>
            public object GetProperty(MSExcel.HEAP_INDEX_COLUMN indx)
            {
                return Properties[(int)(indx - MSExcel._HEAP_INDEX_COLUMN_PROPERTY)].Value;
            }
            /// <summary>
            /// Коструктор объекта со значениями свойств сущности
            /// </summary>
            /// <param name="name">Наименование сущности</param>
            /// <param name="command">Команда для создания сущности</param>
            /// <param name="arProperties">Значения свойств</param>
            public ProxyEntity(string name, MSExcel.COMMAND_ENTITY command, Property[] arProperties)
            {
                m_name = name;

                m_command = command;

                Properties = new Property[arProperties.Length];

                arProperties.CopyTo(Properties, 0);
            }
            /// <summary>
            /// Конструктор - дополнительный (с параметрами)
            ///  индексы назначаются автоматически
            /// </summary>
            /// <param name="arVales">Массив</param>
            public ProxyEntity(object[]arVales)
            {
                m_name = (string)arVales[0];

                m_command = (MSExcel.COMMAND_ENTITY)Enum.Parse(typeof(MSExcel.COMMAND_ENTITY), (string)arVales[1]);
                // ??? почему не константа, где определена (_HEAP_INDEX_COLUMN_PROPERTY)
                Properties = new Property[arVales.Length - 2];

                for (int indx = 0; indx < Properties.Length; indx ++)
                    Properties[indx] = new Property(indx + MSExcel._HEAP_INDEX_COLUMN_PROPERTY
                        , arVales[indx + MSExcel._HEAP_INDEX_COLUMN_PROPERTY]);
            }
        }
        /// <summary>
        /// Создать новый примитив - ящик по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - ящик</returns>
        public static EntityParser.ProxyEntity newBox(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityParser.ProxyEntity pEntityRes;

            string name = string.Empty;
            MSExcel.COMMAND_ENTITY command = MSExcel.COMMAND_ENTITY.UNKNOWN;

            if (TryParseCommandAndNameEntity(format, rEntity, out name, out command) == true)
                // значения для параметров примитива
                switch (format) {
                    case MSExcel.FORMAT.HEAP:
                        pEntityRes = new ProxyEntity(
                            name
                            , command
                            , new ProxyEntity.Property[] {
                                new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.BOX_LAENGTH_X
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.BOX_LAENGTH_X].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.BOX_LAENGTH_Y
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.BOX_LAENGTH_Y].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.BOX_LAENGTH_Z
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.BOX_LAENGTH_Z].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.BOX_PTDISPLACEMENT_X
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.BOX_PTDISPLACEMENT_X].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.BOX_PTDISPLACEMENT_Y
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.BOX_PTDISPLACEMENT_Y].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.BOX_PTDISPLACEMENT_Z
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.BOX_PTDISPLACEMENT_Z].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                            }
                        );

                        //pEntityRes.m_BlockName = blockName;
                        break;
                    case MSExcel.FORMAT.ORDER:
                    default:
                        // соэдать примитив по умолчанию
                        pEntityRes = new ProxyEntity();
                        break;
                }
            else {
                pEntityRes = new ProxyEntity();

                Core.Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Ошибка опрделения имени, типа  сущности..."));
            }

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - конус по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - конус</returns>
        public static EntityParser.ProxyEntity newCone(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityParser.ProxyEntity pEntityRes;

            string name = string.Empty;
            MSExcel.COMMAND_ENTITY command = MSExcel.COMMAND_ENTITY.UNKNOWN;

            if (TryParseCommandAndNameEntity(format, rEntity, out name, out command) == true)
                // значения для параметров примитива
                switch (format) {
                    case MSExcel.FORMAT.HEAP:
                        pEntityRes = new ProxyEntity(
                            name
                            , command
                            , new ProxyEntity.Property[] {
                                new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CONE_HEIGHT
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CONE_HEIGHT].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CONE_ARADIUS_X
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CONE_ARADIUS_X].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CONE_ARADIUS_Y
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CONE_ARADIUS_Y].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CONE_RADIUS_TOP
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CONE_RADIUS_TOP].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CONE_PTDISPLACEMENT_X
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CONE_PTDISPLACEMENT_X].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CONE_PTDISPLACEMENT_Y
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CONE_PTDISPLACEMENT_Y].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CONE_PTDISPLACEMENT_Z
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CONE_PTDISPLACEMENT_Z].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                            }
                        );

                        //pEntityRes.m_BlockName = blockName;
                        break;
                    case MSExcel.FORMAT.ORDER:
                    default:
                        // соэдать примитив по умолчанию
                        pEntityRes = new EntityParser.ProxyEntity();
                        break;
                }
            else {
                pEntityRes = new ProxyEntity();

                Core.Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Ошибка опрделения имени, типа  сущности..."));
            }

            return pEntityRes;
        }
        /// <summary>
        /// Создать примитив - полилинию с бесконечным числом вершин
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - кривая</returns>
        public static EntityParser.ProxyEntity newPolyLine3d(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityParser.ProxyEntity pEntityRes;

            int cntVertex = -1 // количество точек
                , j = -1; // счетчие вершин в цикле
            double[] point3d = new double[3];
            List<double[]> pnts = new List<double[]>() {};

            string name = string.Empty;
            MSExcel.COMMAND_ENTITY command = MSExcel.COMMAND_ENTITY.UNKNOWN;

            if (TryParseCommandAndNameEntity(format, rEntity, out name, out command) == true)
                // значения для параметров примитива
                switch (format) {
                    case MSExcel.FORMAT.HEAP:
                        cntVertex = ((rEntity.Table.Columns.Count - (int)Settings.MSExcel.HEAP_INDEX_COLUMN.POLYLINE_X_START)
                            - (rEntity.Table.Columns.Count - (int)Settings.MSExcel.HEAP_INDEX_COLUMN.POLYLINE_X_START) % 3) / 3;

                        for (j = 0; j < cntVertex; j ++) {
                            if ((!(rEntity[j * 3 + ((int)Settings.MSExcel.HEAP_INDEX_COLUMN.POLYLINE_X_START + 0)] is DBNull))
                                && (!(rEntity[j * 3 + ((int)Settings.MSExcel.HEAP_INDEX_COLUMN.POLYLINE_X_START + 1)] is DBNull))
                                && (!(rEntity[j * 3 + ((int)Settings.MSExcel.HEAP_INDEX_COLUMN.POLYLINE_X_START + 2)] is DBNull)))
                                if ((double.TryParse((string)rEntity[j * 3 + ((int)Settings.MSExcel.HEAP_INDEX_COLUMN.POLYLINE_X_START + 0)], out point3d[0]) == true)
                                    && (double.TryParse((string)rEntity[j * 3 + ((int)Settings.MSExcel.HEAP_INDEX_COLUMN.POLYLINE_X_START + 1)], out point3d[1]) == true)
                                    && (double.TryParse((string)rEntity[j * 3 + ((int)Settings.MSExcel.HEAP_INDEX_COLUMN.POLYLINE_X_START + 2)], out point3d[2]) == true))
                                    pnts.Add(new double[3] { point3d[0], point3d[1], point3d[2] });
                                else
                                    break;
                            else
                                break;
                        }

                        if (pnts.Count > 2) {
                        // соэдать примитив 
                            pEntityRes = new ProxyEntity(
                                name
                                , command
                                , new ProxyEntity.Property[] {
                                    new ProxyEntity.Property (pnts.Count
                                        , pnts)
                                }
                            );

                            //pEntityRes.m_BlockName = blockName;
                        } else {
                            pEntityRes = new ProxyEntity();

                            Core.Logging.DebugCaller(MethodBase.GetCurrentMethod()
                               ,  string.Format(@"Недостаточно точек для создания {0} с именем={1}"
                                    , (string)rEntity[(int)INDEX_KEY.COMMAND]
                                    , (string)rEntity[(int)INDEX_KEY.NAME]
                            ));
                        }
                        break;
                    case MSExcel.FORMAT.ORDER:
                    default:
                        pEntityRes = new ProxyEntity();
                        break;
                }
            else {
                pEntityRes = new ProxyEntity();

                Core.Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Ошибка опрделения имени, типа  сущности..."));
            }

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - окружность по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - окружность</returns>
        public static EntityParser.ProxyEntity newCircle(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityParser.ProxyEntity pEntityRes;

            string name = string.Empty;
            MSExcel.COMMAND_ENTITY command = MSExcel.COMMAND_ENTITY.UNKNOWN;

            if (TryParseCommandAndNameEntity(format, rEntity, out name, out command) == true)
                // значения для параметров примитива
                switch (format) {
                    case MSExcel.FORMAT.HEAP:
                        pEntityRes = new ProxyEntity(
                            name
                            , command
                            , new ProxyEntity.Property[] {
                                new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_CENTER_X
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_CENTER_X].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_CENTER_Y
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_CENTER_Y].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_CENTER_Z
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_CENTER_Z].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_RADIUS
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_RADIUS].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_COLORINDEX
                                    , int.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_COLORINDEX].ToString()
                                        , System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_TICKNESS
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_TICKNESS].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                            }
                        );

                        //pEntityRes.m_BlockName = blockName;
                        break;
                    case MSExcel.FORMAT.ORDER:
                        pEntityRes = new ProxyEntity(
                            name
                            , command
                            , new ProxyEntity.Property[] {
                                new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_CENTER_X
                                    , double.Parse(rEntity[@"CENTER.X"].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_CENTER_Y
                                    , double.Parse(rEntity[@"CENTER.Y"].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_CENTER_Z
                                    , double.Parse(rEntity[@"CENTER.Z"].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_RADIUS
                                    , double.Parse(rEntity[@"Radius"].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_COLORINDEX
                                    , int.Parse(rEntity[@"ColorIndex"].ToString()
                                        , System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_TICKNESS
                                    , double.Parse(rEntity[@"TICKNESS"].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                            }
                        );
                        break;
                    default:
                        // соэдать примитив по умолчанию
                        pEntityRes = new ProxyEntity();
                        break;
                }
            else {
                pEntityRes = new ProxyEntity();

                Core.Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Ошибка опрделения имени, типа  сущности..."));
            }

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - дугу по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - дуга</returns>
        public static EntityParser.ProxyEntity newArc(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityParser.ProxyEntity pEntityRes;

            string name = string.Empty;
            MSExcel.COMMAND_ENTITY command = MSExcel.COMMAND_ENTITY.UNKNOWN;

            if (TryParseCommandAndNameEntity(format, rEntity, out name, out command) == true)
                // значения для параметров примитива
                switch (format) {
                    case MSExcel.FORMAT.HEAP:
                        pEntityRes = new ProxyEntity(
                            name
                            , command
                            , new ProxyEntity.Property[] {
                                new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_CENTER_X
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.ARC_CENTER_X].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_CENTER_Y
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.ARC_CENTER_Y].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_CENTER_Z
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.ARC_CENTER_Z].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_RADIUS
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.ARC_RADIUS].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_ANGLE_START
                                    , (Math.PI / 180) * float.Parse(rEntity[(int)Settings.MSExcel.HEAP_INDEX_COLUMN.ARC_ANGLE_START].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_ANGLE_END
                                    , (Math.PI / 180) * float.Parse(rEntity[(int)Settings.MSExcel.HEAP_INDEX_COLUMN.ARC_ANGLE_END].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_COLORINDEX
                                    , int.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.ARC_COLORINDEX].ToString()
                                        , System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_TICKNESS
                                    , double.Parse(rEntity[(int)MSExcel.HEAP_INDEX_COLUMN.ARC_TICKNESS].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                            }
                        );

                        //pEntityRes.m_BlockName = blockName;
                        break;
                    case MSExcel.FORMAT.ORDER:
                        pEntityRes = new ProxyEntity(
                            name
                            , command
                            , new ProxyEntity.Property[] {
                                new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_CENTER_X
                                    , double.Parse(rEntity[@"CENTER.X"].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_CENTER_Y
                                    , double.Parse(rEntity[@"CENTER.Y"].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_CENTER_Z
                                    , double.Parse(rEntity[@"CENTER.Z"].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_RADIUS
                                    , double.Parse(rEntity[@"Radius"].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_ANGLE_START
                                    , (Math.PI / 180) * float.Parse(rEntity[@"ANGLE.START"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.ARC_ANGLE_END
                                    , (Math.PI / 180) * float.Parse(rEntity[@"ANGLE.END"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_COLORINDEX
                                    , int.Parse(rEntity[@"ColorIndex"].ToString()
                                        , System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture))
                                , new ProxyEntity.Property ((int)MSExcel.HEAP_INDEX_COLUMN.CIRCLE_TICKNESS
                                    , double.Parse(rEntity[@"TICKNESS"].ToString()
                                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture))
                            }
                        );

                        break;
                    default:
                        // соэдать примитив по умолчанию
                        pEntityRes = new ProxyEntity();
                        break;
                }
            else {
                pEntityRes = new ProxyEntity();

                Core.Logging.DebugCaller(MethodBase.GetCurrentMethod(), string.Format(@"Ошибка опрделения имени, типа  сущности..."));
            }

            return pEntityRes;
        }
    }
}
