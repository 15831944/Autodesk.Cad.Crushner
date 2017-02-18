using System;

namespace Autodesk.Cad.Crushner.Assignment
{
    /// <summary>
    /// Уникальный идентификатор примитива на чертеже
    /// </summary>
    public class KEY_ENTITY : Settings.KEY_ENTITY
    {
        /// <summary>
        /// Тип объекта
        /// </summary>
        public Type m_type;
        /// <summary>
        /// Наименование метода для создания сущности
        /// </summary>
        public string m_nameCreateMethod;
        /// <summary>
        /// Наименование типа сложного объекта
        /// </summary>
        public string m_nameSolidType;
        /// <summary>
        /// Конструктор - основной (с параметрами)
        /// </summary>
        /// <param name="blockName">Наименование блока-владельца</param>
        /// <param name="name">Наименование сущности</param>
        public KEY_ENTITY(/*string blockName, */Settings.MSExcel.COMMAND_ENTITY command, int indx, string name)
            : base (/*blockName, */command, indx, name)
        {
            if (!(Valid < 0)) {
                m_type = MSExcel.GetTypeEntity(command);
                m_nameCreateMethod = MSExcel.GetNameCreateMethodEntity(command);
                m_nameSolidType = MSExcel.GetNameSolidTypeEntity(command);
            } else {
                m_type = Type.Missing as Type;
                m_nameCreateMethod =
                m_nameSolidType =
                     string.Empty;
            }
        }
        /// <summary>
        /// Конструктор - основной (с парметрами)
        /// </summary>
        /// <param name="type">Тип сущности</param>
        /// <param name="nameSolidType">Наименование сложного типа(при необходимости)</param>
        /// <param name="nameCreateMethod">Наименование метода для создания сущности</param>
        /// <param name="command">Часть команды для создания сущности</param>
        /// <param name="blockName">Наименование блока-владельца</param>
        /// <param name="indx">Индекс сущности - номер в наименовании, уникальный в пределах блока</param>
        public KEY_ENTITY(string blockName, Settings.MSExcel.COMMAND_ENTITY command, int indx, Type type, string nameSolidType, string nameCreateMethod)
            : base (/*blockName, */command, indx)
        {
            m_type = type;

            m_nameSolidType = nameSolidType;

            m_nameCreateMethod = nameCreateMethod;
        }

        public KEY_ENTITY(MSExcel.MAP_KEY_ENTITY mapKeyEntity, string blockName, int indx)
            : this(blockName, mapKeyEntity.m_command, indx, mapKeyEntity.m_type, mapKeyEntity.m_nameSolidType, mapKeyEntity.m_nameCreateMethod)
        {
        }
        /// <summary>
        /// Оператор срвнения
        /// </summary>
        /// <param name="o1">Объект №1 (слева от операнда) для сравнения</param>
        /// <param name="o2">Объект №2 (справа от операнда) для сравнения</param>
        /// <returns>Результат сравнения</returns>
        public static bool operator==(KEY_ENTITY o1, KEY_ENTITY o2)
        {
            return (((Settings.KEY_ENTITY)o1 == (Settings.KEY_ENTITY)o2) == true)
                && (o1.m_type.Equals(o2.m_type) == true);
        }
        /// <summary>
        /// Оператор срвнения
        /// </summary>
        /// <param name="o1">Объект №1 (слева от операнда) для сравнения</param>
        /// <param name="o2">Объект №2 (справа от операнда) для сравнения</param>
        /// <returns>Результат сравнения</returns>
        public static bool operator !=(KEY_ENTITY o1, KEY_ENTITY o2)
        {
            return (((Settings.KEY_ENTITY)o1 == (Settings.KEY_ENTITY)o2) == false)
                || (o1.m_type.Equals(o2.m_type) == false);
        }
        /// <summary>
        /// Метод для сравнения текущего объекта с объектом, указанным в аргументе
        /// </summary>
        /// <param name="obj">Объект для сравнения</param>
        /// <returns>Результат сравнения</returns>
        public override bool Equals(object obj)
        {
            return this == (KEY_ENTITY)obj;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
