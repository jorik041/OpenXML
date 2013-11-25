using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Framework.Create
{
    public class CellForFooter
    {
        /// <summary>
        /// ячейка
        /// </summary>
        public Cell _Cell { get; private set; }
        /// <summary>
        /// значение
        /// </summary>
        public String Value { get; private set; }

        public CellForFooter(Cell cell, String value)
        {
            _Cell = cell;
            Value = value;
        }
    }
}
