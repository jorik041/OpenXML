using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Framework.Create
{
    /// <summary>
    /// Создание Excel файла
    /// </summary>
    public class Worker
    {
        /// <summary>
        /// путь к папке с шаблонами 
        /// </summary>
        private const String TemplateFolder = "C:\\Templates\\";

        /// <summary>
        /// имя листа шаблона (с которым мы будем работать) 
        /// </summary>
        private const String SheetName = "Лист1";

        /// <summary>
        /// тип документа
        /// </summary>
        private const String FileType = ".xlsx";

        /// <summary>
        /// Папка, для хранения выгруженных файлов
        /// </summary>
        public static String Directory
        {
            get
            {
                const string excelFilesPath = @"C:\xlsx_repository\";
                if (System.IO.Directory.Exists(excelFilesPath) == false)
                {
                    System.IO.Directory.CreateDirectory(excelFilesPath);
                }

                return excelFilesPath;
            }
        }

        public void Export(System.Data.DataTable dataTable, System.Collections.Hashtable hashtable, String templateName)
        {
            var filePath = CreateFile(templateName);

            OpenForRewriteFile(filePath, dataTable, hashtable);

            OpenFile(filePath);
        }

        private String CreateFile(String templateName)
        {
            var templateFelePath = String.Format("{0}{1}{2}", TemplateFolder, templateName, FileType);
            var templateFolderPath = String.Format("{0}{1}", Directory, templateName);
            if (!File.Exists(String.Format("{0}{1}{2}", TemplateFolder, templateName, FileType)))
            {
                throw new Exception(String.Format("Не удалось найти шаблон документа \n\"{0}{1}{2}\"!", TemplateFolder, templateName, FileType));
            }

            //Если в пути шаблона (в templateName) присутствуют папки, то при выгрузке, тоже создаём папки
            var index = (templateFolderPath).LastIndexOf("\\", System.StringComparison.Ordinal);
            if (index > 0)
            {
                var directoryTest = (templateFolderPath).Remove(index, (templateFolderPath).Length - index);
                if (System.IO.Directory.Exists(directoryTest) == false)
                {
                    System.IO.Directory.CreateDirectory(directoryTest);
                }
            }

            var newFilePath = String.Format("{0}_{1}{2}", templateFolderPath, Regex.Replace((DateTime.Now.ToString(CultureInfo.InvariantCulture)), @"[^a-z0-9]+", ""), FileType);
            File.Copy(templateFelePath, newFilePath, true);
            return newFilePath;
        }

        private void OpenForRewriteFile(String filePath, System.Data.DataTable dataTable, System.Collections.Hashtable hashtable)
        {
            Row rowTemplate = null;
            var footer = new List<Footer>();
            var firsIndexFlag = false;
            using (var document = SpreadsheetDocument.Open(filePath, true))
            {
                Sheet sheet;
                try
                {
                    sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().SingleOrDefault(s => s.Name == SheetName);
                }
                catch (Exception ex)
                {
                    throw new Exception(String.Format("Возможно в документе существует два листа с названием \"{0}\"!\n",SheetName), ex);
                }

                if (sheet == null)
                {
                    throw new Exception(String.Format("В шаблоне не найден \"{0}\"!\n",SheetName));
                }

                var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                var rowsForRemove = new List<Row>();
                var fields = new List<Field>();
                foreach (var row in worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>())
                {
                    var celsForRemove = new List<Cell>();
                    foreach (var cell in row.Descendants<Cell>())
                    {
                        if (cell == null)
                        {
                            continue;
                        }

                        var value = GetCellValue(cell, document.WorkbookPart);
                        if (value.IndexOf("DataField:", StringComparison.Ordinal) != -1)
                        {
                            if (!firsIndexFlag)
                            {
                                firsIndexFlag = true;
                                rowTemplate = row;
                            }
                            fields.Add(new Field(Convert.ToUInt32(Regex.Replace(cell.CellReference.Value, @"[^\d]+", ""))
                                , new string(cell.CellReference.Value.ToCharArray().Where(p => !char.IsDigit(p)).ToArray())
                                , value.Replace("DataField:", "")));

                        }

                        if (value.IndexOf("Label:", StringComparison.Ordinal) != -1 && rowTemplate == null)
                        {
                            var labelName = value.Replace("Label:", "").Trim();
                            if (!hashtable.ContainsKey(labelName))
                            {
                                throw new Exception(String.Format("Нет такого лэйбла \"{0}\"", labelName));
                            }
                            cell.CellValue = new CellValue(hashtable[labelName].ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);

                        }

                        if (rowTemplate == null || row.RowIndex <= rowTemplate.RowIndex || String.IsNullOrWhiteSpace(value))
                        {
                            continue;
                        }
                        var item = footer.SingleOrDefault(p => p._Row.RowIndex == row.RowIndex);
                        if (item == null)
                        {
                            footer.Add(new Footer(row, cell, value.IndexOf("Label:", StringComparison.Ordinal) != -1 ? hashtable[value.Replace("Label:", "").Trim()].ToString() : value));
                        }
                        else
                        {
                            item.AddMoreCell(cell, value.IndexOf("Label:", StringComparison.Ordinal) != -1 ? hashtable[value.Replace("Label:", "").Trim()].ToString() : value);
                        }
                        celsForRemove.Add(cell);
                    }

                    foreach (var cell in celsForRemove)
                    {
                        cell.Remove();
                    }

                    if (rowTemplate != null && row.RowIndex != rowTemplate.RowIndex)
                    {
                        rowsForRemove.Add(row);
                    }
                }

                if (rowTemplate == null || rowTemplate.RowIndex == null || rowTemplate.RowIndex < 0)
                {
                    throw new Exception("Не удалось найти ни одного поля, для заполнения!");
                }

                foreach (var row in rowsForRemove)
                {
                    row.Remove();
                }

                var index = rowTemplate.RowIndex;
                foreach (var row in from System.Data.DataRow item in dataTable.Rows select CreateRow(rowTemplate, index, item, fields))
                {
                    sheetData.InsertBefore(row, rowTemplate);
                    index++;
                }

                foreach (var newRow in footer.Select(item => CreateLabel(item, (UInt32)dataTable.Rows.Count)))
                {
                    sheetData.InsertBefore(newRow, rowTemplate);
                }

                rowTemplate.Remove();
            }
        }

        private Row CreateLabel(Footer item, uint count)
        {
            var row = item._Row;
            row.RowIndex = new UInt32Value(item._Row.RowIndex + (count - 1));
            foreach (var cell in item.Cells)
            {
                cell._Cell.CellReference = new StringValue(cell._Cell.CellReference.Value.Replace(Regex.Replace(cell._Cell.CellReference.Value, @"[^\d]+", ""), row.RowIndex.ToString()));
                cell._Cell.CellValue = new CellValue(cell.Value);
                cell._Cell.DataType = new EnumValue<CellValues>(CellValues.String);
                row.Append(cell._Cell);
            }
            return row;
        }

        private Row CreateRow(Row rowTemplate, uint index, System.Data.DataRow item, List<Field> fields)
        {
            var newRow = (Row)rowTemplate.Clone();
            newRow.RowIndex = new UInt32Value(index);

            foreach (var cell in newRow.Elements<Cell>())
            {
                cell.CellReference = new StringValue(cell.CellReference.Value.Replace(Regex.Replace(cell.CellReference.Value, @"[^\d]+", ""), index.ToString(CultureInfo.InvariantCulture)));
                foreach (var fil in fields.Where(fil => cell.CellReference == fil.Column + index))
                {
                    cell.CellValue = new CellValue(item[fil._Field].ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }
            return newRow;
        }


        private string GetCellValue(Cell cell, WorkbookPart wbPart)
        {
            var value = cell.InnerText;

            if (cell.DataType == null)
            {
                return value;
            }
            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:

                    var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    if (stringTable != null)
                    {
                        value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }
                    break;
            }

            return value;
        }

        private void OpenFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new Exception(String.Format("Не удалось найти файл \"{0}\"!", filePath));
            }

            var process = Process.Start(filePath);
            if (process != null)
            {
                process.WaitForExit();
            }
        }
    }
}
