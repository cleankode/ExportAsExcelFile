using Excel.Export.Core;
using OpenXML.Excel.Export.Employees;
using System;
using System.IO;

namespace Client
{
    class Program
    {
        static void Main(string[] args)
        {
            var defaultColor = Console.ForegroundColor;

            var fileName = Path.Combine(Environment.CurrentDirectory, "Employees.xlsx");
            IExcelFileGenerator report = new ExcelFileGenerator();

            report.GenerateExcelFile(fileName);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Excel file has created!");
            Console.ForegroundColor = defaultColor;

            Console.WriteLine("Press [Enter] to exit.");
            Console.ReadLine();
        }
    }
}
