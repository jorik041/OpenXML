using System;
using System.Collections.Generic;
using System.Data;

namespace OpenXmlPrj
{
    class Program
    {
        static void Main(string[] args)
        {
            //заполняем тестовыми данными
            var myData = new List<DataForTest>
            {
                new DataForTest("a1","b1","c1"),
                new DataForTest("a2","b2","c2"),
                new DataForTest("a3","b3","c3"),
                new DataForTest("a4","b4","c4"),
                new DataForTest("a5","b5","c5")
            };

            var ex = new Converter.ConvertToDataTable();
            //ex.ExcelTableLines(myData) - конвертируем наши данные в DataTable
            //ex.ExcelTableHeader(myData.Count) - формируем данные для Label
            //template - указываем название нашего файла  - шаблона
            new Framework.Create.Worker().Export(ex.ExcelTableLines(myData), ex.ExcelTableHeader(myData.Count), "template");

            Console.WriteLine("Excel File Has Created!\nFor Read Data From Excel, press any key!");
            Console.ReadKey();
            //"C:\\Loading\\ReadMePlease.xlsx" - путь к файлу, с которого будем считывать данные (возвращяет нам DataTable)
            var dt = new Framework.Load.Worker().ReadFile("C:\\Loading\\ReadMePlease.xlsx");
            var myDataFromExcel = new List<DataForTest>();
            //Заполняем наш объект, считанными данными из DataTable
            foreach (DataRow item in dt.Rows)
            {
                myDataFromExcel.Add(new DataForTest(item));
            }

            Console.WriteLine("---------- Data ---------------------");
            //Выводим считанные данные
            foreach (var line in myDataFromExcel)
            {
                Console.WriteLine("{0} | {1} | {2}", line.A, line.B, line.C);
            }

            Console.WriteLine("Done. Press any key, for exit!");
            Console.ReadKey();
        }
    }
}
