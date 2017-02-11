using System;

namespace Autodesk.Cad.Crushner.Settings
{
    /// <summary>
    /// Уникальный идентификатор примитива на чертеже
    /// </summary>
    public class KEY_ENTITY
    {
        /// <summary>
        /// Разделитель в наименовании сущности (тип-индекс)
        /// </summary>
        protected static Char s_chNameDelimeter = '-';
        /// <summary>
        /// Часть команды для создания сущности
        /// </summary>
        public Settings.MSExcel.COMMAND_ENTITY m_command;
        /// <summary>
        /// Индекс сущности - номер в наименовании, уникальный в пределах блока
        /// </summary>
        public int m_index;
        /// <summary>
        /// Уникальное наименование со специальной сигнатурой (состоит из 2-х частей)
        ///  1 - COMMAND_ENTITY
        ///  2 - 3-х значный цифровой индекс
        /// </summary>
        public string Name {
            get {
                if (!(Valid < 0))
                    if (Valid == 0)
                        return string.Format(@"{0}{2}{1:000}", m_command.ToString(), m_index, s_chNameDelimeter);
                    else
                        return string.Format(@"{0}", m_command.ToString());
                else
                    return string.Empty;
            }
        }

        public int Valid;
        /// <summary>
        /// Наименование блока, к которому принадлежит сущность
        /// </summary>
        public string m_BlockName;
        /// <summary>
        /// Конструктор - основной (с параметрами)
        /// </summary>
        /// <param name="blockName">Наименование блока-владельца</param>
        /// <param name="nameEntity">Наименование сущности</param>
        public KEY_ENTITY(string blockName, string nameEntity)
        {
            MSExcel.COMMAND_ENTITY command = MSExcel.COMMAND_ENTITY.UNKNOWN;
            string[] names = nameEntity.Split(s_chNameDelimeter);

            m_BlockName = string.Empty;

            if (names.Length == 2)
                if (Enum.TryParse(names[0], out command) == true) {
                    if (Int32.TryParse(names[1], out m_index) == true) {
                        m_command = command;

                        Valid = 0;
                    } else
                        Valid = 1;

                    m_BlockName = blockName;
                } else
                    Valid = -1;
            else
                Valid = -2;
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
        public KEY_ENTITY(Settings.MSExcel.COMMAND_ENTITY command, string blockName, int indx)
        {
            m_command = command;

            m_index = indx;

            m_BlockName = blockName;
        }
        /// <summary>
        /// Оператор срвнения
        /// </summary>
        /// <param name="o1">Объект №1 (слева от операнда) для сравнения</param>
        /// <param name="o2">Объект №2 (справа от операнда) для сравнения</param>
        /// <returns>Результат сравнения</returns>
        public static bool operator==(KEY_ENTITY o1, KEY_ENTITY o2)
        {
            return (o1.m_command.Equals(o2.m_command) == true)
                && (o1.m_index.Equals(o2.m_index) == true);
        }
        /// <summary>
        /// Оператор срвнения
        /// </summary>
        /// <param name="o1">Объект №1 (слева от операнда) для сравнения</param>
        /// <param name="o2">Объект №2 (справа от операнда) для сравнения</param>
        /// <returns>Результат сравнения</returns>
        public static bool operator !=(KEY_ENTITY o1, KEY_ENTITY o2)
        {
            return (o1.m_command.Equals(o2.m_command) == false)
                || (o1.m_index.Equals(o2.m_index) == false);
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
