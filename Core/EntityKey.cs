using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.Interop.Common;

namespace Autodesk.Cad.Crushner.Core
{
    /// <summary>
    /// Перечисление - типы примитивов
    /// </summary>
    public enum COMMAND_ENTITY : short
    {
        UNKNOWN = -1
        , CIRCLE, ARC, LINE, PLINE3, CONE, BOX
            , COUNT
    }
    /// <summary>
    /// Уникалбный идентификатор примитива на чертеже
    /// </summary>
    public struct KEY_ENTITY
    {
        private static Char s_chNameDelimeter = '-';

        public COMMAND_ENTITY m_command;

        public int m_index;
        /// <summary>
        /// Уникальное наименование со специальной сигнатурой (состоит из 2-х частей)
        ///  1 - COMMAND_ENTITY
        ///  2 - 3-х значный цифровой индекс
        /// </summary>
        public string Name {
            get {
                if (!(m_command == COMMAND_ENTITY.UNKNOWN))
                    if (!(m_index < 0))
                        return string.Format(@"{0}{2}{1:000}", m_command.ToString(), m_index, s_chNameDelimeter);
                    else
                        return string.Format(@"{0}", m_command.ToString());
                else
                    return string.Empty;
            }
        }
        /// <summary>
        /// Тип объекта
        /// </summary>
        public Type m_type;

        public string m_nameCreateMethod;

        public string m_nameSolidType;

        public string m_BlockName;

        public KEY_ENTITY(string blockName, string name)
        {
            string[] names = name.Split(s_chNameDelimeter);

            if ((Enum.TryParse(names[0], out m_command) == true)
                && (Int32.TryParse(names[1], out m_index) == true)) {
                m_type = MSExcel.GetTypeEntity(m_command);
                m_nameCreateMethod = MSExcel.GetNameCreateMethodEntity(m_command);
                m_nameSolidType = MSExcel.GetNameSolidTypeEntity(m_command);
                m_BlockName = blockName;
            } else {
                m_command = COMMAND_ENTITY.UNKNOWN;
                m_index = -1;
                m_type = Type.Missing as Type;
                m_nameCreateMethod =
                m_nameSolidType =
                     string.Empty;
                m_BlockName = string.Empty;
            }
        }

        public KEY_ENTITY(Type type, string nameSolidType, string nameCreateMethod, COMMAND_ENTITY command, string blockName, int indx)
        {
            m_nameCreateMethod = nameCreateMethod;

            m_type = type;

            m_nameSolidType = nameSolidType;

            m_nameCreateMethod = nameCreateMethod;

            m_command = command;

            m_BlockName = blockName;

            m_index = indx;
        }

        public static bool operator==(KEY_ENTITY o1, KEY_ENTITY o2)
        {
            return (o1.m_command.Equals(o2.m_command) == true)
                && (o1.m_index.Equals(o2.m_index) == true)
                && (o1.m_type.Equals(o2.m_type) == true);
        }

        public static bool operator !=(KEY_ENTITY o1, KEY_ENTITY o2)
        {
            return (o1.m_command.Equals(o2.m_command) == false)
                || (o1.m_index.Equals(o2.m_index) == false)
                || (o1.m_type.Equals(o2.m_type) == false);
        }

        public override bool Equals(object obj)
        {
            return this == (KEY_ENTITY)obj;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }

    public class Collection
    {
        public static readonly List<KEY_ENTITY> s_MappingKeyEntity = new List<KEY_ENTITY>() {
            new KEY_ENTITY () { m_command = COMMAND_ENTITY.CIRCLE, m_index = -1, m_type = typeof(Circle), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new KEY_ENTITY () { m_command = COMMAND_ENTITY.ARC, m_index = -1, m_type = typeof(Arc), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new KEY_ENTITY () { m_command = COMMAND_ENTITY.LINE, m_index = -1, m_type = typeof(Line), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new KEY_ENTITY () { m_command = COMMAND_ENTITY.PLINE3, m_index = -1, m_type = typeof(Polyline3d), m_nameSolidType = string.Empty, m_nameCreateMethod = string.Empty }
            , new KEY_ENTITY () { m_command = COMMAND_ENTITY.CONE, m_index = -1, m_type = typeof(Solid3d), m_nameSolidType = @"Frustum", m_nameCreateMethod = @"CreateFrustum" }
            , new KEY_ENTITY () { m_command = COMMAND_ENTITY.BOX, m_index = -1, m_type = typeof(Solid3d), m_nameSolidType = @"Box", m_nameCreateMethod = @"CreateBox" }
            ,
        };

        public static string GetNameSolidTypeEntity(COMMAND_ENTITY command)
        {
            string nameSolidTypeRes = string.Empty;

            KEY_ENTITY keyEntity =
                s_MappingKeyEntity.Find(item => { return item.m_command == command; });

            if (!(keyEntity == null))
                nameSolidTypeRes = keyEntity.m_nameSolidType;
            else
                ;

            return nameSolidTypeRes;
        }

        public static string GetNameCreateMethodEntity(COMMAND_ENTITY command)
        {
            string nameMethodRes = string.Empty;

            KEY_ENTITY keyEntity =
                s_MappingKeyEntity.Find(item => { return item.m_command == command; });

            if (!(keyEntity == null))
                nameMethodRes = keyEntity.m_nameCreateMethod;
            else
                ;

            return nameMethodRes;
        }

        public static Type GetTypeEntity(COMMAND_ENTITY command)
        {
            Type typeRes = Type.Missing as Type;

            KEY_ENTITY keyEntity = 
                s_MappingKeyEntity.Find(item => { return item.m_command == command; });

            if (!(keyEntity == null))
                typeRes = keyEntity.m_type;
            else
                ;

            return typeRes;
        }

        public static MSExcel.DictionaryBlock s_dictBlock = new MSExcel.DictionaryBlock();

        public static void Clear()
        {
            s_dictBlock.Clear();
        }
    }

}
