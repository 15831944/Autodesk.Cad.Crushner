using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using static Autodesk.Cad.Crushner.Settings.Collection;

namespace Autodesk.Cad.Crushner.Core
{
    public partial class EntityCtor
    {
        public struct ProxyEntity
        {
            public Entity m_entity;

            public Point3d m_ptDisplacement;

            //public string m_BlockName;

            public ProxyEntity(Entity entity)
            {
                m_entity = entity;

                m_ptDisplacement = Point3d.Origin;

                //m_BlockName = string.Empty;
            }

            public void SetPoint3dDisplacement(double[] values)
            {
                m_ptDisplacement = new Point3d(values);
            }
        }

        //public class ProxyEntity : Entity
        //{
        //    public HProxyEntity(IntPtr unManagedObj, bool bAutoDelete)
        //        : base(unManagedObj, bAutoDelete)
        //    {
        //    }

        //    public static implicit operator Solid3d(HProxyEntity pEntity)
        //    {
        //        return new Solid3d(pEntity.XData.UnmanagedObject, pEntity.XData.AutoDelete);
        //    }

        //    public static implicit operator HProxyEntity(Solid3d oSolid3d)
        //    {
        //        return new HProxyEntity(oSolid3d.UnmanagedObject, oSolid3d.AutoDelete);
        //    }
        //}        
        /// <summary>
        /// Создать новый примитив - ящик по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - ящик</returns>
        public static EntityCtor.ProxyEntity newBox(object[] arEntity)
        {
            EntityCtor.ProxyEntity pEntityRes;
            double lAlongX = -1F, lAlongY = -1F, lAlongZ = -1;
            double []ptDisplacement = new double[3];

            pEntityRes = new ProxyEntity (new Solid3d());

            // значения для параметров примитива
            switch (format)
            {
                case MSExcel.FORMAT.HEAP:
                    lAlongX = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.BOX_LAENGTH_X].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    lAlongY = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.BOX_LAENGTH_Y].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    lAlongZ = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.BOX_LAENGTH_Z].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    ptDisplacement[0] = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.BOX_PTDISPLACEMENT_X].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    ptDisplacement[1] = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.BOX_PTDISPLACEMENT_Y].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    ptDisplacement[2] = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.BOX_PTDISPLACEMENT_Z].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);

                    ((Solid3d)pEntityRes.m_entity).CreateBox(lAlongX, lAlongY, lAlongZ);
                    pEntityRes.SetPoint3dDisplacement (ptDisplacement);

                    //pEntityRes.m_BlockName = blockName;
                    break;
                case MSExcel.FORMAT.ORDER:
                    break;
                default:
                    break;
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
        public static EntityCtor.ProxyEntity newCone(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;

            double height = -1F
                , rAlongX = -1F, rAlongY = -1
                , rTop = -1;
            double[] ptDisplacement = new double[3];

            KEY_ENTITY keyEntity =
                Collection.s_MappingKeyEntity.Find(item => {
                    return item.m_command.Equals(Enum.Parse(typeof(COMMAND_ENTITY), (string)rEntity[1])) == true;
                });
            ConstructorInfo coneCtor = keyEntity.m_type.GetConstructor(Type.EmptyTypes);
            MethodInfo methodCreate = keyEntity.m_type.GetMethod(keyEntity.m_nameCreateMethod);

            pEntityRes = new EntityCtor.ProxyEntity();
            pEntityRes.m_entity = null;
            pEntityRes.m_ptDisplacement = Point3d.Origin;

            // значения для параметров примитива
            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    height = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CONE_HEIGHT].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    rAlongX = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CONE_ARADIUS_X].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    rAlongY = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CONE_ARADIUS_Y].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    rTop = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CONE_RADIUS_TOP].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    ptDisplacement[0] = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CONE_PTDISPLACEMENT_X].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    ptDisplacement[1] = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CONE_PTDISPLACEMENT_Y].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    ptDisplacement[2] = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CONE_PTDISPLACEMENT_Z].ToString()
                        , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);

                    //pEntityRes = new ProxyEntity (new Solid3d());
                    pEntityRes.m_entity = coneCtor.Invoke(new object[] { }) as Solid3d;
                    //(pEntityRes.m_entity as Solid3d).CreateFrustum(height, rAlongX, rAlongY, rTop);
                    methodCreate.Invoke(pEntityRes.m_entity, new object[] { height, rAlongX, rAlongY, rTop }); // CreateFrustum
                    pEntityRes.SetPoint3dDisplacement(ptDisplacement);

                    //pEntityRes.m_BlockName = blockName;
                    break;
                case MSExcel.FORMAT.ORDER:
                    break;
                default:
                    break;
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
        public static EntityCtor.ProxyEntity newPolyLine3d(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;

            int cntVertex = -1 // количество точек
                , j = -1; // счетчие вершин в цикле
            double[] point3d = new double[3];
            Point3dCollection pnts = new Point3dCollection() {};

            KEY_ENTITY keyEntity =
                Collection.s_MappingKeyEntity.Find(item => {
                    return item.m_command.Equals(Enum.Parse(typeof(COMMAND_ENTITY), (string)rEntity[1])) == true;
                });
            ConstructorInfo coneCtor = keyEntity.m_type.GetConstructor(new Type[] {
                typeof(Poly3dType)
                , typeof(Point3dCollection)
                , typeof(bool)
            });

            pEntityRes = new ProxyEntity();
            pEntityRes.m_entity = null;
            pEntityRes.m_ptDisplacement = Point3d.Origin;

            // значения для параметров примитива
            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    cntVertex = ((rEntity.Table.Columns.Count - (int)Settings.Collection.HEAP_INDEX_COLUMN.POLYLINE_X_START)
                        - (rEntity.Table.Columns.Count - (int)Settings.Collection.HEAP_INDEX_COLUMN.POLYLINE_X_START) % 3) / 3;

                    for (j = 0; j < cntVertex; j ++) {
                        if ((!(rEntity[j * 3 + ((int)Settings.Collection.HEAP_INDEX_COLUMN.POLYLINE_X_START + 0)] is DBNull))
                            && (!(rEntity[j * 3 + ((int)Settings.Collection.HEAP_INDEX_COLUMN.POLYLINE_X_START + 1)] is DBNull))
                            && (!(rEntity[j * 3 + ((int)Settings.Collection.HEAP_INDEX_COLUMN.POLYLINE_X_START + 2)] is DBNull)))
                            if ((double.TryParse((string)rEntity[j * 3 + ((int)Settings.Collection.HEAP_INDEX_COLUMN.POLYLINE_X_START + 0)], out point3d[0]) == true)
                                && (double.TryParse((string)rEntity[j * 3 + ((int)Settings.Collection.HEAP_INDEX_COLUMN.POLYLINE_X_START + 1)], out point3d[1]) == true)
                                && (double.TryParse((string)rEntity[j * 3 + ((int)Settings.Collection.HEAP_INDEX_COLUMN.POLYLINE_X_START + 2)], out point3d[2]) == true))
                                pnts.Add(new Point3d(point3d));
                            else
                                break;
                        else
                            break;
                    }

                    if (pnts.Count > 2) {
                    // соэдать примитив 
                        pEntityRes.m_entity =
                            //new EntityParser.ProxyEntity (new Polyline3d(Poly3dType.SimplePoly, pnts, false))
                            coneCtor.Invoke(new object[] { Poly3dType.SimplePoly, pnts, false })
                            as Entity;

                        //pEntityRes.m_BlockName = blockName;
                    } else {
                        Logging.AcEditorWriteMessage(string.Format(@"Недостаточно точек для создания {0} с именем={1}"
                            , (string)rEntity[1]
                            , (string)rEntity[0]
                        ));
                    }
                    break;
                case MSExcel.FORMAT.ORDER:
                    break;
                default:
                    break;
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
        public static EntityCtor.ProxyEntity newCircle(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;
            // соэдать примитив 
            pEntityRes = new ProxyEntity (new Circle());
            // значения для параметров примитива
            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    (pEntityRes.m_entity as Circle).Center = new Point3d(
                        double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CIRCLE_CENTER_X].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CIRCLE_CENTER_Y].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CIRCLE_CENTER_Z].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (pEntityRes.m_entity as Circle).Radius = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CIRCLE_RADIUS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (pEntityRes.m_entity as Circle).ColorIndex = int.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CIRCLE_COLORINDEX].ToString());
                    (pEntityRes.m_entity as Circle).Thickness = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.CIRCLE_TICKNESS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);

                    //pEntityRes.m_BlockName = blockName;
                    break;
                case MSExcel.FORMAT.ORDER:
                    (pEntityRes.m_entity as Circle).Center = new Point3d(
                        double.Parse(rEntity[@"CENTER.X"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"CENTER.Y"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"CENTER.Z"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (pEntityRes.m_entity as Circle).Radius = double.Parse(rEntity[@"Radius"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (pEntityRes.m_entity as Circle).ColorIndex = int.Parse(rEntity[@"ColorIndex"].ToString());
                    (pEntityRes.m_entity as Circle).Thickness = double.Parse(rEntity[@"TICKNESS"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                default:
                    break;
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
        public static EntityCtor.ProxyEntity newArc(DataRow rEntity, MSExcel.FORMAT format/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;
            // соэдать примитив 
            pEntityRes = new ProxyEntity (new Arc());
            // значения для параметров примитива
            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    (pEntityRes.m_entity as Arc).Center = new Point3d(
                        double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ARC_CENTER_X].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ARC_CENTER_Y].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ARC_CENTER_Z].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (pEntityRes.m_entity as Arc).Radius = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ARC_RADIUS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (pEntityRes.m_entity as Arc).StartAngle = (Math.PI / 180) * float.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ARC_ANGLE_START].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (pEntityRes.m_entity as Arc).EndAngle = (Math.PI / 180) * float.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ARC_ANGLE_END].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (pEntityRes.m_entity as Arc).ColorIndex = int.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ARC_COLORINDEX].ToString());
                    (pEntityRes.m_entity as Arc).Thickness = double.Parse(rEntity[(int)Settings.Collection.HEAP_INDEX_COLUMN.ARC_TICKNESS].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);

                    //pEntityRes.m_BlockName = blockName;
                    break;
                case MSExcel.FORMAT.ORDER:
                    (pEntityRes.m_entity as Arc).Center = new Point3d(
                        double.Parse(rEntity[@"CENTER.X"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"CENTER.Y"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                        , double.Parse(rEntity[@"CENTER.Z"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                    (pEntityRes.m_entity as Arc).Radius = double.Parse(rEntity[@"Radius"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (pEntityRes.m_entity as Arc).StartAngle = (Math.PI / 180) * float.Parse(rEntity[@"ANGLE.START"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (pEntityRes.m_entity as Arc).EndAngle = (Math.PI / 180) * float.Parse(rEntity[@"ANGLE.END"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    (pEntityRes.m_entity as Arc).ColorIndex = int.Parse(rEntity[@"ColorIndex"].ToString());
                    (pEntityRes.m_entity as Arc).Thickness = double.Parse(rEntity[@"TICKNESS"].ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                default:
                    break;
            }

            return pEntityRes;
        }

        public static object[] boxToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair, MSExcel.FORMAT format)
        {
            object[] rowRes = null;

            switch (format)
            {
                case MSExcel.FORMAT.HEAP:
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    break;
            }

            return rowRes;
        }

        public static object[] coneToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair, MSExcel.FORMAT format)
        {
            object[] rowRes = null;

            double height = -1F
                , rAlongX = -1F, rAlongY = -1
                , rTop = -1;
            Point3d pt3dPosition
                , pt3dMin = Point3d.Origin, pt3dMax = Point3d.Origin;

            pt3dMin = (pair.Value.m_entity as Solid3d).GeometricExtents.MinPoint;
            pt3dMax = (pair.Value.m_entity as Solid3d).GeometricExtents.MaxPoint;
            rAlongX = pt3dMax.X - pt3dMin.X / 2;
            rAlongY = pt3dMax.Y - pt3dMin.Y / 2;
            height = pt3dMax.Z - pt3dMin.Z;
            //rTop = ???

            switch (format)
            {
                case MSExcel.FORMAT.HEAP:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0}", pair.Key.m_command.ToString()) //!!! COMMAND_ENTITY
                        , string.Format(@"{0:0.0}", height) //CONE_HEIGHT
                        , string.Format(@"{0:0.0}", rAlongX) //CONE_ARADIUS_X
                        , string.Format(@"{0:0.0}", rAlongY) //CONE_ARADIUS_X
                        , string.Format(@"{0:0}", (pair.Value.m_entity as Solid3d).GetField(@"rTop")) //Radius
                        //??? координаты размещения
                    };
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    break;
            }

            return rowRes;
        }

        public static object[] polyLine3dToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair, MSExcel.FORMAT format)
        {
            object[] rowRes = null;

            int cntVertex = -1 // количество вершин полилинии
                , j = -1;
            Point3d point3d;

            switch (format)
            {
                case MSExcel.FORMAT.HEAP:
                    cntVertex = (int)(pair.Value.m_entity as Polyline3d).Length;

                    rowRes = new object[(int)Settings.Collection.HEAP_INDEX_COLUMN.POLYLINE_X_START + cntVertex];

                    rowRes[0] = string.Format(@"{0}", pair.Key.Name); //NAME
                    rowRes[1] = string.Format(@"{0}", pair.Key.m_command.ToString()); //!!! COMMAND_ENTITY

                    for (j = 0; j < cntVertex; j++) {
                        point3d = (pair.Value.m_entity as Polyline3d).GetPointAtParameter((double)j);

                        rowRes[j * 3 + ((int)Settings.Collection.HEAP_INDEX_COLUMN.POLYLINE_X_START + 0)] = string.Format(@"{0:0.0}", point3d.X); //X
                        rowRes[j * 3 + ((int)Settings.Collection.HEAP_INDEX_COLUMN.POLYLINE_X_START + 1)] = string.Format(@"{0:0.0}", point3d.Y); //Y
                        rowRes[j * 3 + ((int)Settings.Collection.HEAP_INDEX_COLUMN.POLYLINE_X_START + 2)] = string.Format(@"{0:0.0}", point3d.Z); //Z
                    }
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    break;
            }

            return rowRes;
        }

        public static object[] circleToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair, MSExcel.FORMAT format)
        {
            object[] rowRes = null;

            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0}", pair.Key.m_command.ToString()) //!!! COMMAND_ENTITY
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Circle).Center.X) //CENTER.X
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Circle).Center.Y) //CENTER.Y
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Circle).Center.Z) //CENTER.Z
                        , string.Format(@"{0:0}", (pair.Value.m_entity as Circle).Radius) //Radius
                        , string.Format(@"{0}", (pair.Value.m_entity as Circle).ColorIndex) //ColorIndex
                        , string.Format(@"{0:0.000}", (pair.Value.m_entity as Circle).Thickness) //Tickness
                    };
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Circle).Center.X) //CENTER.X
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Circle).Center.Y) //CENTER.Y
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Circle).Center.Z) //CENTER.Z
                        , string.Format(@"{0:0}", (pair.Value.m_entity as Circle).Radius) //Radius
                        , string.Format(@"{0}", (pair.Value.m_entity as Circle).ColorIndex) //ColorIndex
                        , string.Format(@"{0:0.000}", (pair.Value.m_entity as Circle).Thickness) //Tickness
                    };
                    break;
            }

            return rowRes;
        }

        public static object[] arcToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair, MSExcel.FORMAT format)
        {
            object[] rowRes = null;

            switch (format) {
                case MSExcel.FORMAT.HEAP:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0}", pair.Key.m_command.ToString()) //!!! COMMAND_ENTITY
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Arc).Center.X) //CENTER.X
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Arc).Center.Y) //CENTER.Y
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Arc).Center.Z) //CENTER.Z
                        , string.Format(@"{0:0}", (pair.Value.m_entity as Arc).Radius) //Radius
                        , string.Format(@"{0:0}", (180 / Math.PI) * (pair.Value.m_entity as Arc).StartAngle) //START.ANGLE
                        , string.Format(@"{0:0}", (180 / Math.PI) * (pair.Value.m_entity as Arc).EndAngle) //END.ANGLE
                        , string.Format(@"{0}", (pair.Value.m_entity as Arc).ColorIndex) //ColorIndex
                        , string.Format(@"{0:0.000}", (pair.Value.m_entity as Arc).Thickness) //Tickness
                    };
                    break;
                case MSExcel.FORMAT.ORDER:
                default:
                    rowRes = new object[] {
                        string.Format(@"{0}", pair.Key.Name) //NAME
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Arc).Center.X) //CENTER.X
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Arc).Center.Y) //CENTER.Y
                        , string.Format(@"{0:0.0}", (pair.Value.m_entity as Arc).Center.Z) //CENTER.Z
                        , string.Format(@"{0:0}", (pair.Value.m_entity as Arc).Radius) //Radius
                        , string.Format(@"{0:0}", (180 / Math.PI) * (pair.Value.m_entity as Arc).StartAngle) //START.ANGLE
                        , string.Format(@"{0:0}", (180 / Math.PI) * (pair.Value.m_entity as Arc).EndAngle) //END.ANGLE
                        , string.Format(@"{0}", (pair.Value.m_entity as Arc).ColorIndex) //ColorIndex
                        , string.Format(@"{0:0.000}", (pair.Value.m_entity as Arc).Thickness) //Tickness
                    };
                    break;
            }

            return rowRes;
        }
    }
}
