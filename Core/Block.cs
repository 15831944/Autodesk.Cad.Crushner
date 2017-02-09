using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop.Common;
using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Autodesk.Cad.Crushner.Core
{
    partial class MSExcel
    {
        public class BLOCK
        {
            public struct PLACEMENT
            {
                public Point3d m_point3d;

                public override string ToString()
                {
                    return string.Format(@"координаты [{0}, {1}, {2}]"
                        , m_point3d.X
                        , m_point3d.Y
                        , m_point3d.Z
                    );
                }

                public PLACEMENT(Point3d pt3d)
                {
                    m_point3d = pt3d;
                }
            }

            public List<PLACEMENT> m_ListReference;

            public Dictionary<KEY_ENTITY, EntityParser.ProxyEntity> m_dictEntity;

            public BLOCK()
            {
                m_dictEntity = new Dictionary<KEY_ENTITY, EntityParser.ProxyEntity>();
            }

            public BLOCK(ExcelWorksheet ews) : this()
            {
            }

            public void Add(EntityParser.ProxyEntity pEntity)
            {
                m_dictEntity.Add(GetKeyEntity(
                        pEntity.m_entity
                        , pEntity.m_entity is Solid3d ? ((Acad3DSolid)pEntity.m_entity).SolidType : string.Empty
                    )
                    , pEntity
                );
            }

            public int GetCount(object entity)
            {
                int iRes = 0;

                iRes = m_dictEntity.Keys.Where(item => (
                    (item.m_type == entity.GetType())
                    && (item.m_type.Equals(typeof(Solid3d)) == true ?
                        (item.m_nameSolidType.Equals(((Acad3DSolid)entity).SolidType) == true) :
                        true) == true)).Count();

                return iRes;
            }

            public KEY_ENTITY GetKeyEntity(object entity, string nameSolidType)
            {
                KEY_ENTITY keyEntityRes = s_MappingKeyEntity.Find(item => (
                    (item.m_type.Equals(entity.GetType()) == true)
                    && (item.m_nameSolidType == nameSolidType))
                );

                keyEntityRes.m_index = GetCount(entity);

                return keyEntityRes;
            }
        }

        public class DictionaryBlock : Dictionary<string, BLOCK>
        {
            private void addItem (string blockName)
            {
                this.Add(blockName, new BLOCK());
                this[blockName].m_ListReference = new List<BLOCK.PLACEMENT>();
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

                    this[blockName].m_ListReference.Add(new BLOCK.PLACEMENT(new Point3d(pt3dvalues)));
                } else
                    ; // ошибка добавления ссылки на блок
            }

            public void AddEntity(string blockName, string name, EntityParser.ProxyEntity pEntity)
            {
                if (this.ContainsKey(blockName) == false)
                    this.Add(blockName, new BLOCK());
                else
                    ;

                this[blockName].m_dictEntity.Add(new KEY_ENTITY(blockName, name), pEntity);
            }
            /// <summary>
            /// Добавить примитив в список, подготовленный для экспорта (что есть на чертеже)
            /// </summary>
            /// <param name="pEntity">Примитив для добавления</param>
            /// <return>Признак наличия примитива в составе блока</return>
            public bool AddToExport(EntityParser.ProxyEntity pEntity)
            {
                bool bRes = this[pEntity.m_entity.BlockName].m_dictEntity.Values.Contains(pEntity);

                if (bRes == false)
                    this[pEntity.m_entity.BlockName].m_dictEntity.Add(s_dictBlock[pEntity.m_entity.BlockName].GetKeyEntity(
                            pEntity
                            , pEntity is Solid3d ? (pEntity.m_entity as Acad3DSolid).SolidType : string.Empty
                        )
                        , pEntity
                    );
                else
                    ;

                return bRes;
            }

            public KEY_ENTITY GetKeyEntity(EntityParser.ProxyEntity pEntity)
            {
                return this[pEntity.m_entity.BlockName].m_dictEntity.FirstOrDefault(x => (x.Value.m_entity as DBObject).ObjectId == (pEntity.m_entity as DBObject).ObjectId).Key;
            }
        }
    }
}
