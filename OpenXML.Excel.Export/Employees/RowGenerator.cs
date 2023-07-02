using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel.Export.Core;
using System.Collections.Generic;
using System.Globalization;

namespace OpenXML.Excel.Export.Employees
{
    public class Generate
    {
        internal static OpenXmlElement[] EmployeeRow(Employee employee)
        {
            return new OpenXmlElement[]
            {
                CellHelper.ConstructCell(employee.Id.ToString(), CellValues.Number, 1),
                CellHelper.ConstructCell(employee.Name, CellValues.String, 1),
                CellHelper.ConstructCell(employee.DoB.ToString("yyyy/MM/dd"), CellValues.String, 1),
                CellHelper.ConstructCell(employee.YearlySalary.ToString(CultureInfo.InvariantCulture), CellValues.Number, 1)
            };
        }

        internal static OpenXmlElement[] AddressesRow(Employee employee)
        {
            return new OpenXmlElement[]
            {
                CellHelper.ConstructCell(employee.Address.Id.ToString(), CellValues.Number, 1),
                CellHelper.ConstructCell(employee.Id.ToString(), CellValues.Number, 1),
                CellHelper.ConstructCell(employee.Address.Line1, CellValues.String, 1),
                CellHelper.ConstructCell(employee.Address.Line2, CellValues.String, 1),
                CellHelper.ConstructCell(employee.Address.City, CellValues.String, 1),
                CellHelper.ConstructCell(employee.Address.State, CellValues.String, 1),
                CellHelper.ConstructCell(employee.Address.Line2, CellValues.String, 1)
            };
        }

        internal class HeaderSettingFor
        {
            private static IDictionary<OpenXmlElement, OpenXmlElement> _employees;
            private static IDictionary<OpenXmlElement, OpenXmlElement> _addresses;

            /// <summary>
            /// Sets up Employees sheet column headers
            /// </summary>
            /// <returns>Column names with their setting</returns>
            internal static IDictionary<OpenXmlElement, OpenXmlElement> Employees
            {
                get
                {
                    if (_employees == null)
                    {
                        _employees = new Dictionary<OpenXmlElement, OpenXmlElement>
                    {
                        { CellHelper.ConstructCell("Id", CellValues.String, 2), new Column { Min = 1, Max = 1, Width = 4, CustomWidth = true} },
                        { CellHelper.ConstructCell("Name", CellValues.String, 2),  new Column { Min = 2, Max = 3, Width = 15, CustomWidth = true } },
                        { CellHelper.ConstructCell("Birth Date", CellValues.String, 2), new Column { Min = 2, Max = 3, Width = 15, CustomWidth = true } },
                        { CellHelper.ConstructCell("Salary/year", CellValues.String, 2), new Column { Min = 4, Max = 4, Width = 12, CustomWidth = true, BestFit=true } }
                    };
                    }

                    return _employees;
                }
            }

            /// <summary>
            /// Sets up Addresses sheet column headers
            /// </summary>
            /// <returns>Column names with their setting</returns>
            internal static IDictionary<OpenXmlElement, OpenXmlElement> Addresses
            {
                get
                {
                    if (_addresses == null)
                    {
                        _addresses = new Dictionary<OpenXmlElement, OpenXmlElement>
                    {
                        { CellHelper.ConstructCell("Id", CellValues.String, 2),new Column {Min = 1,Max = 1,Width = 4,CustomWidth = true } },
                        { CellHelper.ConstructCell("EmployeeID", CellValues.String, 2),new Column {Min = 1,Max = 1,Width = 4,CustomWidth = true} },
                        { CellHelper.ConstructCell("Line1", CellValues.String, 2),new Column {Min = 10,Max = 20,Width = 15,CustomWidth = true} },
                        { CellHelper.ConstructCell("Line2", CellValues.String, 2),new Column {Min = 4,Max = 4,Width = 8,CustomWidth = true} },
                        { CellHelper.ConstructCell("City", CellValues.String, 2),new Column{ Min = 10, Max = 20, Width = 15,CustomWidth = true} },
                        { CellHelper.ConstructCell("State", CellValues.String, 2),new Column { Min = 2, Max = 2,Width = 5,CustomWidth = true} },
                        { CellHelper.ConstructCell("Zip Code", CellValues.String, 2),new Column {Min = 5,Max = 11,Width = 10,CustomWidth = true } }
                    };
                    }

                    return _addresses;
                }
            }
        }
    }
}
