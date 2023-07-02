using Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel.Export.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace OpenXML.Excel.Export.Employees
{
    public class ExcelFileGenerator : IExcelFileGenerator
    {
        private readonly IEmployeeRepository _employeeRepository;

        public ExcelFileGenerator()
        {
            _employeeRepository = new EmployeeRepository(); //FIXME - for production code, you shouldn't inistantiate here of course but instead - use some sort of IOC container
        }

        public void GenerateExcelFile(string fileName)
        {
            try
            {
                //FIXME - when fileName isn't provided, use default one.
                var employees = _employeeRepository.GetEmployees();

                using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    stylePart.Stylesheet = CellHelper.GenerateStylesheet();
                    stylePart.Stylesheet.Save();
                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                    CreateEmployeesTab(workbookPart, sheets, employees);
                    CreateAddressesTab(workbookPart, sheets, employees);
                }
            }
            catch(Exception exc)
            {
                Debug.WriteLine(exc);
            }
        }

        private void CreateAddressesTab(WorkbookPart workbookPart, Sheets sheets, IEnumerable<Employee> employees)
        {
            WorksheetPart worksheetpart2 = workbookPart.AddNewPart<WorksheetPart>();
            Worksheet worksheet2 = new Worksheet();
            SheetData sheetdata2 = new SheetData();

            // Setting up columns & constructing addresses header
            worksheet2.AppendChild(new Columns(Generate.HeaderSettingFor.Addresses.Values));
            Row addressRow = new Row();
            addressRow.Append(Generate.HeaderSettingFor.Addresses.Keys);

            Sheet addressesSheet = new Sheet { Id = workbookPart.GetIdOfPart(worksheetpart2), SheetId = 2, Name = "Addresses" };
            sheets.Append(addressesSheet);

            // Insert the header
            sheetdata2.AppendChild(addressRow);

            // Inserting address rows
            foreach (var employee in employees)
            {
                addressRow = new Row();
                addressRow.Append(Generate.AddressesRow(employee));
                sheetdata2.AppendChild(addressRow);
            }

            worksheet2.AppendChild(sheetdata2);
            worksheetpart2.Worksheet = worksheet2;
        }

        private void CreateEmployeesTab(WorkbookPart workbookPart, Sheets sheets, IEnumerable<Employee> employees)
        {
            //employees sheet
            WorksheetPart worksheetpart1 = workbookPart.AddNewPart<WorksheetPart>();
            Worksheet worksheet1 = new Worksheet();
            SheetData sheetdata1 = new SheetData();

            // Setting up columns & construct headers
            worksheet1.AppendChild(new Columns(Generate.HeaderSettingFor.Employees.Values));
            Row employeeRow = new Row();
            employeeRow.Append(Generate.HeaderSettingFor.Employees.Keys);

            // create Employees sheet
            Sheet employeesSheet = new Sheet { Id = workbookPart.GetIdOfPart(worksheetpart1), SheetId = 1, Name = "Employees" };
            sheets.Append(employeesSheet);

            // Insert the header row to the Sheet Data
            sheetdata1.AppendChild(employeeRow);

            // Inserting each employee
            foreach (var employee in employees)
            {
                employeeRow = new Row();
                employeeRow.Append(Generate.EmployeeRow(employee));
                sheetdata1.AppendChild(employeeRow);
            }

            worksheet1.AppendChild(sheetdata1);
            worksheetpart1.Worksheet = worksheet1;
        }
    }
}