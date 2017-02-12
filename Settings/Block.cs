using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Autodesk.Cad.Crushner.Settings
{
    partial class MSExcel
    {
        /// <summary>
        /// Координаты точки для размещения блока
        /// </summary>
        public struct POINT3D
        {
            public double X;

            public double Y;

            public double Z;

            public override string ToString()
            {
                return string.Format(@"координаты [{0}, {1}, {2}]"
                    , X
                    , Y
                    , Z
                );
            }

            public POINT3D(double[] values)
            {
                if (values.Length == 3) {
                    X = values[0];
                    Y = values[1];
                    Z = values[2];
                }
                else
                    throw new Exception(@"PLACEMENT невозможно определить: кол-во точек не равно 3...");
            }

            public double[] Values { get { return new double[] { X, Y, Z }; } }
        }

        public class BLOCK
        {
            /// <summary>
            /// Список ссылок(копий) блока
            /// </summary>
            public List<POINT3D> m_ListReference;

            public Dictionary<KEY_ENTITY, EntityParser.ProxyEntity> m_dictEntityParser;

            public BLOCK()
            {
                m_dictEntityParser = new Dictionary<KEY_ENTITY, EntityParser.ProxyEntity>();
            }
            /// <summary>
            /// Конструктор - дополнительный (с параметром)
            ///  из перечгя объектов на одном листе файла конфигурации (книги VS Excel)
            /// </summary>
            /// <param name="ews">Лист файла конфигурации (книги MS Excel)</param>
            public BLOCK(ExcelWorksheet ews) : this()
            {
            }

            public void Add(KEY_ENTITY key, EntityParser.ProxyEntity entity)
            {
                m_dictEntityParser.Add(key, entity);
            }
        }

        public class DictionaryBlock : Dictionary<string, BLOCK>
        {
            protected void addItem (string blockName)
            {
                this.Add(blockName, new BLOCK());
                this[blockName].m_ListReference = new List<POINT3D>();
            }

            public void AddReference(DataRow rReference)
            {
                string blockName = string.Empty;
                int iColumn = -1;
                double[] pt3dvalues = new double[3];

                for (iColumn = 0; iColumn < rReference.ItemArray.Length; iColumn++)
                    if (!(rReference[iColumn] is DBNull))
                        switch (iColumn) {
                            case 0:
                                blockName = (string)rReference[iColumn];
                                break;
                            default:
                                if (double.TryParse((string)rReference[iColumn], out pt3dvalues[iColumn - 1]) == false)
                                    iColumn = rReference.ItemArray.Length;
                                else
                                    ;
                                break;
                        }
                    else
                    // предупреждение - в ячейке 'Null'
                        ;

                if (!(iColumn < rReference.ItemArray.Length)) {
                    if (this.ContainsKey(blockName) == false)
                        addItem(blockName);
                    else
                        ;

                    this[blockName].m_ListReference.Add(new POINT3D(pt3dvalues));
                } else
                    ; // ошибка добавления ссылки на блок
            }

            public void AddEntity(string blockName, string name, EntityParser.ProxyEntity pEntity)
            {
                if (this.ContainsKey(blockName) == false)
                    this.Add(blockName, new BLOCK());
                else
                    ;

                this[blockName].Add(new KEY_ENTITY(blockName, name), pEntity);
            }
        }
    }
}
