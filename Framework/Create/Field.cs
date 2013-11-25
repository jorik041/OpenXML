using System;

namespace Framework.Create
{
    public class Field
    {
        /// <summary>
        /// Индекс строки
        /// </summary>
        public uint Row { get; private set; }
        /// <summary>
        /// координаты колонки
        /// </summary>
        public String Column { get; private set; }
        /// <summary>
        /// название колонки, выводимых данных
        /// </summary>
        public String _Field { get; private set; }

        public Field(uint row, String column, String field)
        {
            Row = row;
            Column = column;
            _Field = field;
        }
    }
}
