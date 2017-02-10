using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Settings
{
    /// <summary>
    /// Базовый класс - коллекция для хранения данных из файла конфигурации
    /// </summary>
    public class Collection
    {
        /// <summary>
        /// Перечисление - индексы столбцов на листе книги MS Excel в формате 'HEAP'
        /// </summary>
        public enum HEAP_INDEX_COLUMN
        {
            CIRCLE_CENTER_X = 2, CIRCLE_CENTER_Y, CIRCLE_CENTER_Z, CIRCLE_RADIUS, CIRCLE_COLORINDEX, CIRCLE_TICKNESS
            , ARC_CENTER_X = 2, ARC_CENTER_Y, ARC_CENTER_Z, ARC_RADIUS, ARC_ANGLE_START, ARC_ANGLE_END, ARC_COLORINDEX, ARC_TICKNESS
            , LINE_START_X = 2, LINE_START_Y, LINE_START_Z, LINE_END_X, LINE_END_Y, LINE_END_Z, LINE_COLORINDEX, LINE_TICKNESS
            , POLYLINE_X_START = 2
            , CONE_HEIGHT = 2, CONE_ARADIUS_X, CONE_ARADIUS_Y, CONE_RADIUS_TOP, CONE_PTDISPLACEMENT_X, CONE_PTDISPLACEMENT_Y, CONE_PTDISPLACEMENT_Z
            , BOX_LAENGTH_X = 2, BOX_LAENGTH_Y, BOX_LAENGTH_Z, BOX_PTDISPLACEMENT_X, BOX_PTDISPLACEMENT_Y, BOX_PTDISPLACEMENT_Z
            , ALINE_START_X = 2, ALINE_START_Y, ALINE_START_Z, ALINE_LENGTH, ALINE_COLORINDEX, ALINE_TICKNESS
            , RLINE_NAME_ENTITY_RELATIVE = 2, RLINE_LENGTH, RLINE_COLORINDEX, RLINE_TICKNESS
            ,
        }
    }
}
