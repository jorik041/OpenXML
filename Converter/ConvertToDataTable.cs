using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using Interfaces;

namespace Converter
{
    public class ConvertToDataTable
    {
        public DataTable ExcelTableLines(IEnumerable<IDataForTest> lines)
        {
            var dt = CreateTable();
            foreach (var line in lines)
            {
                var row = dt.NewRow();
                row["AAA"] = line.A;
                row["BBB"] = line.B;
                row["CCC"] = line.C;
                dt.Rows.Add(row);
            }
            return dt;
        }

        public Hashtable ExcelTableHeader(Int32 count)
        {
            var head = new Dictionary<String, String> { { "Date", DateTime.Today.Date.ToShortDateString() }, { "Count", count.ToString() } };
            return new Hashtable(head);
        }

        private DataTable CreateTable()
        {
            var dt = new DataTable("ExelTable");
            var col = new DataColumn { DataType = typeof(String), ColumnName = "AAA" };
            dt.Columns.Add(col);
            col = new DataColumn { DataType = typeof(String), ColumnName = "BBB" };
            dt.Columns.Add(col);
            col = new DataColumn { DataType = typeof(String), ColumnName = "CCC" };
            dt.Columns.Add(col);
            return dt;
        }
    }
}
