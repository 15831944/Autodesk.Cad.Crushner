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
    public partial class EntityCtor
    {
        /// <summary>
        /// Класс для хранения промежуточных сведений о сущности для отображения
        ///  между созданием сущности и отображением на чертеже
        /// </summary>
        public struct ProxyEntity
        {
            /// <summary>
            /// Объект сущности для отображения на чертеже
            /// </summary>
            public Entity m_entity;
            /// <summary>
            /// Точка переноса сущности после создания в Point3d.Origin
            /// </summary>
            public Point3d m_ptDisplacement;

            //public string m_BlockName;
            /// <summary>
            /// Конструктор - основной (с параметрами)
            /// </summary>
            /// <param name="entity">Объект сущности</param>
            public ProxyEntity(Entity entity)
            {
                m_entity = entity;

                m_ptDisplacement = Point3d.Origin;

                //m_BlockName = string.Empty;
            }
            /// <summary>
            /// Усиановить значение точки переноса
            /// </summary>
            /// <param name="values">Значения по осям системы координат для точки переноса</param>
            public void SetPoint3dDisplacement(double[] values)
            {
                m_ptDisplacement = new Point3d(values);
            }
        }
        /// <summary>
        /// Создать новый примитив - ящик по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - ящик</returns>
        public static EntityCtor.ProxyEntity newBox(Settings.EntityParser.ProxyEntity entity/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;
            double lAlongX = -1F, lAlongY = -1F, lAlongZ = -1;
            double []ptDisplacement = new double[3];

            pEntityRes = new ProxyEntity (new Solid3d());

            // значения для параметров примитива
            lAlongX = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.BOX_LAENGTH_X).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            lAlongY = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.BOX_LAENGTH_Y).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            lAlongZ = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.BOX_LAENGTH_Z).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            ptDisplacement[0] = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.BOX_PTDISPLACEMENT_X).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            ptDisplacement[1] = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.BOX_PTDISPLACEMENT_Y).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            ptDisplacement[2] = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.BOX_PTDISPLACEMENT_Z).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);

            ((Solid3d)pEntityRes.m_entity).CreateBox(lAlongX, lAlongY, lAlongZ);
            pEntityRes.SetPoint3dDisplacement (ptDisplacement);

            //pEntityRes.m_BlockName = blockName;

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - конус по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - конус</returns>
        public static EntityCtor.ProxyEntity newCone(Settings.EntityParser.ProxyEntity entity/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;

            double height = -1F
                , rAlongX = -1F, rAlongY = -1
                , rTop = -1;
            double[] ptDisplacement = new double[3];

            MSExcel.MAP_KEY_ENTITY mapKeyEntity =
                MSExcel.s_MappingKeyEntity.Find(item => {
                    return item.m_command.Equals(entity.m_command) == true;
                });
            ConstructorInfo coneCtor = mapKeyEntity.m_type.GetConstructor(Type.EmptyTypes);
            MethodInfo methodCreate = mapKeyEntity.m_type.GetMethod(mapKeyEntity.m_nameCreateMethod);

            pEntityRes = new EntityCtor.ProxyEntity();
            pEntityRes.m_entity = null;
            pEntityRes.m_ptDisplacement = Point3d.Origin;

            // значения для параметров примитива
            height = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CONE_HEIGHT).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            rAlongX = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CONE_ARADIUS_X).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            rAlongY = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CONE_ARADIUS_Y).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            rTop = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CONE_RADIUS_TOP).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            ptDisplacement[0] = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CONE_PTDISPLACEMENT_X).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            ptDisplacement[1] = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CONE_PTDISPLACEMENT_Y).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            ptDisplacement[2] = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CONE_PTDISPLACEMENT_Z).ToString()
                , System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);

            //pEntityRes = new ProxyEntity (new Solid3d());
            pEntityRes.m_entity = coneCtor.Invoke(new object[] { }) as Solid3d;
            //(pEntityRes.m_entity as Solid3d).CreateFrustum(height, rAlongX, rAlongY, rTop);
            methodCreate.Invoke(pEntityRes.m_entity, new object[] { height, rAlongX, rAlongY, rTop }); // CreateFrustum
            pEntityRes.SetPoint3dDisplacement(ptDisplacement);

            //pEntityRes.m_BlockName = blockName;

            return pEntityRes;
        }
        /// <summary>
        /// Создать примитив - полилинию с бесконечным числом вершин
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - кривая</returns>
        public static EntityCtor.ProxyEntity newPolyLine3d(Settings.EntityParser.ProxyEntity entity/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;

            int cntVertex = -1 // количество точек
                , j = -1; // счетчие вершин в цикле
            double[] point3d = new double[3];
            Point3dCollection pnts = new Point3dCollection() {};
            PolylineVertex3d vex3d;

            MSExcel.MAP_KEY_ENTITY mapKeyEntity =
                MSExcel.s_MappingKeyEntity.Find(item => {
                    return item.m_command.Equals(entity.m_command) == true;
                });
            ConstructorInfo pline3dCtor = mapKeyEntity.m_type.GetConstructor (
                new Type[] {
                    // вариант №1
                    typeof(Poly3dType)
                    , typeof(Point3dCollection)
                    , typeof(bool)
                    //// вариант №2
                    //
                }
            );

            pEntityRes = new ProxyEntity();
            pEntityRes.m_entity = null;
            pEntityRes.m_ptDisplacement = Point3d.Origin;

            // ??? значения для параметров примитива
            cntVertex = entity.Properties[0].Index;
            foreach (double[]pt3d in entity.Properties[0].Value as List<double[]>)
                pnts.Add(new Point3d (pt3d));

            if (pnts.Count > 2) {
            // соэдать примитив 
                pEntityRes.m_entity =
                    // вариант №1
                    pline3dCtor.Invoke(new object[] { Poly3dType.SimplePoly, pnts, false })
                        //// вариант №2
                        //pline3dCtor.Invoke(new object[] { })
                        as Entity;

                //// для варианта №2
                //foreach (Point3d pt in pnts) {
                //    vex3d = new PolylineVertex3d(pt);
                //    (pEntityRes.m_entity as Polyline3d).AppendVertex(vex3d);
                //}
                //(pEntityRes.m_entity as Polyline3d).Close();

                //pEntityRes.m_BlockName = blockName;
            } else
                Logging.AcEditorWriteMessage(string.Format(@"Недостаточно точек для создания {0} с именем={1}"
                    , entity.m_command
                    , entity.m_name
                ));

            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - окружность по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы со значениями параметров примитива</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - окружность</returns>
        public static EntityCtor.ProxyEntity newCircle(Settings.EntityParser.ProxyEntity entity/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;
            // соэдать примитив 
            pEntityRes = new ProxyEntity (new Circle());
            // значения для параметров примитива
            (pEntityRes.m_entity as Circle).Center = new Point3d(
                double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CIRCLE_CENTER_X).ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                , double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CIRCLE_CENTER_Y).ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                , double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CIRCLE_CENTER_Z).ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
            (pEntityRes.m_entity as Circle).Radius = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CIRCLE_RADIUS).ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            (pEntityRes.m_entity as Circle).ColorIndex = int.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CIRCLE_COLORINDEX).ToString());
            (pEntityRes.m_entity as Circle).Thickness = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.CIRCLE_TICKNESS).ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);

            //pEntityRes.m_BlockName = blockName;
            return pEntityRes;
        }
        /// <summary>
        /// Создать новый примитив - дугу по значениям параметров из строки таблицы
        /// </summary>
        /// <param name="rEntity">Строка таблицы</param>
        /// <param name="format">Формат файла конфигурации из которого была импортирована таблица</param>
        /// <param name="blockName">Наимнование блока (только при формате 'HEAP')</param>
        /// <returns>Объект примитива - дуга</returns>
        public static EntityCtor.ProxyEntity newArc(Settings.EntityParser.ProxyEntity entity/*, string blockName*/)
        {
            EntityCtor.ProxyEntity pEntityRes;
            // соэдать примитив 
            pEntityRes = new ProxyEntity (new Arc());
            // значения для параметров примитива
            (pEntityRes.m_entity as Arc).Center = new Point3d(
                double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ARC_CENTER_X).ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                , double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ARC_CENTER_Y).ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)
                , double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ARC_CENTER_Z).ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
            (pEntityRes.m_entity as Arc).Radius = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ARC_RADIUS).ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            (pEntityRes.m_entity as Arc).StartAngle = (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ARC_ANGLE_START);
            (pEntityRes.m_entity as Arc).EndAngle = (double)entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ARC_ANGLE_END);
            (pEntityRes.m_entity as Arc).ColorIndex = int.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ARC_COLORINDEX).ToString());
            (pEntityRes.m_entity as Arc).Thickness = double.Parse(entity.GetProperty(Settings.MSExcel.HEAP_INDEX_COLUMN.ARC_TICKNESS).ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);

            //pEntityRes.m_BlockName = blockName;

            return pEntityRes;
        }

        public static Settings.EntityParser.ProxyEntity boxToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes = new Settings.EntityParser.ProxyEntity();

            return rowRes;
        }

        public static Settings.EntityParser.ProxyEntity coneToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes;

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

            rowRes = new Settings.EntityParser.ProxyEntity (new object[] {
                string.Format(@"{0}", pair.Key.Name) //NAME
                , string.Format(@"{0}", pair.Key.Command.ToString()) //!!! COMMAND_ENTITY
                , string.Format(@"{0:0.0}", height) //CONE_HEIGHT
                , string.Format(@"{0:0.0}", rAlongX) //CONE_ARADIUS_X
                , string.Format(@"{0:0.0}", rAlongY) //CONE_ARADIUS_X
                , string.Format(@"{0:0}", (pair.Value.m_entity as Solid3d).GetField(@"rTop")) //Radius
                //??? координаты размещения
            });

            return rowRes;
        }

        public static Settings.EntityParser.ProxyEntity polyLine3dToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes
                = new Settings.EntityParser.ProxyEntity()
                ;

            int cntVertex = -1 // количество вершин полилинии
                , j = -1;
            Point3d point3d;

            cntVertex = (int)(pair.Value.m_entity as Polyline3d).Length;

            //rowRes = new object[(int)Settings.MSExcel.HEAP_INDEX_COLUMN.POLYLINE_X_START + cntVertex];

            //rowRes[0] = string.Format(@"{0}", pair.Key.Name); //NAME
            //rowRes[1] = string.Format(@"{0}", pair.Key.m_command.ToString()); //!!! COMMAND_ENTITY

            //for (j = 0; j < cntVertex; j++) {
            //    point3d = (pair.Value.m_entity as Polyline3d).GetPointAtParameter((double)j);

            //    rowRes[j * 3 + ((int)Settings.MSExcel.HEAP_INDEX_COLUMN.POLYLINE_X_START + 0)] = string.Format(@"{0:0.0}", point3d.X); //X
            //    rowRes[j * 3 + ((int)Settings.MSExcel.HEAP_INDEX_COLUMN.POLYLINE_X_START + 1)] = string.Format(@"{0:0.0}", point3d.Y); //Y
            //    rowRes[j * 3 + ((int)Settings.MSExcel.HEAP_INDEX_COLUMN.POLYLINE_X_START + 2)] = string.Format(@"{0:0.0}", point3d.Z); //Z
            //}

            return rowRes;
        }

        public static Settings.EntityParser.ProxyEntity circleToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes;

            rowRes = new Settings.EntityParser.ProxyEntity(new object[] {
                string.Format(@"{0}", pair.Key.Name) //NAME
                , string.Format(@"{0}", pair.Key.Command.ToString()) //!!! COMMAND_ENTITY
                , string.Format(@"{0:0.0}", (pair.Value.m_entity as Circle).Center.X) //CENTER.X
                , string.Format(@"{0:0.0}", (pair.Value.m_entity as Circle).Center.Y) //CENTER.Y
                , string.Format(@"{0:0.0}", (pair.Value.m_entity as Circle).Center.Z) //CENTER.Z
                , string.Format(@"{0:0}", (pair.Value.m_entity as Circle).Radius) //Radius
                , string.Format(@"{0}", (pair.Value.m_entity as Circle).ColorIndex) //ColorIndex
                , string.Format(@"{0:0.000}", (pair.Value.m_entity as Circle).Thickness) //Tickness
            });

            return rowRes;
        }

        public static Settings.EntityParser.ProxyEntity arcToDataRow(KeyValuePair<KEY_ENTITY, EntityCtor.ProxyEntity> pair)
        {
            Settings.EntityParser.ProxyEntity rowRes;

            rowRes = new Settings.EntityParser.ProxyEntity (new object[] {
                string.Format(@"{0}", pair.Key.Name) //NAME
                , string.Format(@"{0}", pair.Key.Command.ToString()) //!!! COMMAND_ENTITY
                , string.Format(@"{0:0.0}", (pair.Value.m_entity as Arc).Center.X) //CENTER.X
                , string.Format(@"{0:0.0}", (pair.Value.m_entity as Arc).Center.Y) //CENTER.Y
                , string.Format(@"{0:0.0}", (pair.Value.m_entity as Arc).Center.Z) //CENTER.Z
                , string.Format(@"{0:0}", (pair.Value.m_entity as Arc).Radius) //Radius
                , string.Format(@"{0:0}", (180 / Math.PI) * (pair.Value.m_entity as Arc).StartAngle) //START.ANGLE
                , string.Format(@"{0:0}", (180 / Math.PI) * (pair.Value.m_entity as Arc).EndAngle) //END.ANGLE
                , string.Format(@"{0}", (pair.Value.m_entity as Arc).ColorIndex) //ColorIndex
                , string.Format(@"{0:0.000}", (pair.Value.m_entity as Arc).Thickness) //Tickness
            });

            return rowRes;
        }
    }
}
