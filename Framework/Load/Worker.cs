using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Framework.Load
{
    /// <summary>
    /// Загрузка данных из Excel
    /// </summary>
    public class Worker
    {
        /// <summary>
        /// имя листа (откуда будем читать данные) 
        /// </summary>
        private const String SheetName = "Лист1";

        /// <summary>
        /// Подавать только файлы в формате .xlsx
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public System.Data.DataTable ReadFile(String path)
        {
            CheckFile(path);
            return OpenDocumentForRead(path);
        }

        private System.Data.DataTable OpenDocumentForRead(string path)
        {
            System.Data.DataTable data = null;
            using (var document = SpreadsheetDocument.Open(path, false))
            {
                Sheet sheet;
                try
                {
                    sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().SingleOrDefault(s => s.Name == SheetName);
                }
                catch (Exception ex)
                {
                    throw new Exception(String.Format("Возможно в документе существует два листа с названием \"{0}\"!\n", SheetName), ex);
                }

                if (sheet == null)
                {
                    throw new Exception(String.Format("В файле не найден \"{0}\"!\n", SheetName));
                }

                var relationshipId = sheet.Id.Value;
                var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                var firstRow = true;
                var columsNames = new List<ColumnName>();
                foreach (Row row in sheetData.Elements<Row>())
                {
                    if (firstRow)
                    {
                        columsNames.AddRange(GetNames(row, document.WorkbookPart));
                        data = GetTable(columsNames);
                        firstRow = false;
                        continue;
                    }

                    var item = data.NewRow();
                    foreach (var line in columsNames)
                    {
                        var coordinates = String.Format("{0}{1}", line.Liter, row.RowIndex);
                        var cc = row.Elements<Cell>().SingleOrDefault(p => p.CellReference == coordinates);
                        if (cc == null)
                        {
                            throw new Exception(String.Format("Не удалось найти ячейку \"{0}\"!", coordinates));
                        }
                        item[line.Name.Trim()] = GetVal(cc, document.WorkbookPart);

                    }
                    data.Rows.Add(item);
                }
            }

            return data;
        }

        private System.Data.DataTable GetTable(IEnumerable<ColumnName> columsNames)
        {
            var teb = new System.Data.DataTable("ExelTable");

            foreach (var col in columsNames.Select(columnName => new System.Data.DataColumn { DataType = typeof(String), ColumnName = columnName.Name.Trim() }))
            {
                teb.Columns.Add(col);
            }

            return teb;
        }

        private IEnumerable<ColumnName> GetNames(Row row, WorkbookPart wbPart)
        {
            return (from cell in row.Elements<Cell>()
                    where cell != null
                    let
                        text = GetVal(cell, wbPart)
                    where !String.IsNullOrWhiteSpace(text)
                    select
                    new ColumnName(text, Regex.Replace(cell.CellReference.Value, @"[\0-9]", ""))).ToList();
        }

        private string GetVal(Cell cell, WorkbookPart wbPart)
        {
            string value = cell.InnerText;

            if (cell.DataType == null)
            {
                return value;
            }
            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:

                    var stringTable =
                        wbPart.GetPartsOfType<SharedStringTablePart>()
                            .FirstOrDefault();

                    if (stringTable != null)
                    {
                        value =
                            stringTable.SharedStringTable
                                .ElementAt(int.Parse(value)).InnerText;
                    }
                    break;
            }

            return value;
        }

        private void CheckFile(String path)
        {
            if (String.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                throw new Exception(String.Format("Такого файла \"{0}\", не существует!", path));
            }
        }
    }
}
