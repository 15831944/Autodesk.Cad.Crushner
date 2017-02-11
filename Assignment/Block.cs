using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Autodesk.Cad.Crushner.Assignment
{
    partial class MSExcel
    {
        public class BLOCK : Settings.MSExcel.BLOCK
        {
            public Dictionary<KEY_ENTITY, EntityCtor.ProxyEntity> m_dictEntityCtor;

            public BLOCK() : base ()
            {
            }

            public void Add(EntityCtor.ProxyEntity pEntity)
            {
                m_dictEntityCtor.Add(GetKeyEntity(
                        pEntity.m_entity
                        , pEntity.m_entity is Solid3d ? ((Acad3DSolid)pEntity.m_entity).SolidType : string.Empty
                    )
                    , pEntity
                );
            }

            public int GetCount(object entity)
            {
                int iRes = 0;

                iRes = m_dictEntityCtor.Keys.Where(item => (
                    (item.m_type == entity.GetType())
                    && (item.m_type.Equals(typeof(Solid3d)) == true ?
                        (item.m_nameSolidType.Equals(((Acad3DSolid)entity).SolidType) == true) :
                        true) == true)).Count();

                return iRes;
            }

            public KEY_ENTITY GetKeyEntity(object entity, string nameSolidType)
            {
                KEY_ENTITY keyEntityRes;

                MAP_KEY_ENTITY mapKeyEntity = s_MappingKeyEntity.Find(item => (
                    (item.m_type.Equals(entity.GetType()) == true)
                    && (item.m_nameSolidType == nameSolidType))
                );

                keyEntityRes = new KEY_ENTITY(mapKeyEntity
                    , ((entity is Solid3d) == true) ? (entity as Solid3d).BlockName : ((entity is Entity) == true) ? (entity as Entity).BlockName : string.Empty
                    , GetCount(entity));

                return keyEntityRes;
            }
        }

        public class DictionaryBlock : Dictionary<string, BLOCK> //Settings.MSExcel.DictionaryBlock
        {
            private DictionaryBlock() : base ()
            {
            }

            public DictionaryBlock(Settings.MSExcel.DictionaryBlock dictBlock) : this()
            {
                Create(dictBlock);
            }

            public void Create(Settings.MSExcel.DictionaryBlock dictBlock)
            {
                foreach (string blockName in dictBlock.Keys) {
                    foreach (KeyValuePair<Settings.KEY_ENTITY, Settings.EntityParser.ProxyEntity> pair in dictBlock[blockName].m_dictEntityParser)
                        AddEntity(blockName, pair.Key.Name, dictDelegateMethodeEntity[pair.Key.m_command].newEntity(pair.Value));

                    //foreach (Settings.MSExcel.BLOCK.PLACEMENT placement in dictBlock[blockName].m_ListReference)
                    //    AddEntity(blockName, pair.Key.Name, dictDelegateMethodeEntity[pair.Key.m_command].newEntity(pair.Value));
                }
            }

            public void AddEntity(string blockName, string name, EntityCtor.ProxyEntity pEntity)
            {
                if (this.ContainsKey(blockName) == false) {
                    this.Add(blockName, new BLOCK());

                    this[blockName].m_dictEntityCtor = new Dictionary<KEY_ENTITY, EntityCtor.ProxyEntity>();
                } else
                    ;

                this[blockName].m_dictEntityCtor.Add(new KEY_ENTITY(blockName, name), pEntity);
            }
            /// <summary>
            /// Добавить примитив в список, подготовленный для экспорта (что есть на чертеже)
            /// </summary>
            /// <param name="pEntity">Примитив для добавления</param>
            /// <return>Признак наличия примитива в составе блока</return>
            public bool AddToExport(EntityCtor.ProxyEntity pEntity)
            {
                bool bRes = (this[pEntity.m_entity.BlockName] as BLOCK).m_dictEntityCtor.Values.Contains(pEntity);

                if (bRes == false)
                    (this[pEntity.m_entity.BlockName] as BLOCK).m_dictEntityCtor.Add((s_dictBlock[pEntity.m_entity.BlockName] as BLOCK).GetKeyEntity(
                            pEntity
                            , pEntity is Solid3d ? (pEntity.m_entity as Acad3DSolid).SolidType : string.Empty
                        )
                        , pEntity
                    );
                else
                    ;

                return bRes;
            }

            public KEY_ENTITY GetKeyEntity(EntityCtor.ProxyEntity pEntity)
            {
                return (this[pEntity.m_entity.BlockName] as BLOCK).m_dictEntityCtor.FirstOrDefault(x => (x.Value.m_entity as DBObject).ObjectId == (pEntity.m_entity as DBObject).ObjectId).Key;
            }
        }
    }
}
