# Export C# object collection to Excel file

This is a sample project to generate Excel file from C# objection collection. 
In this porject, we have an Employee model that have basic info and an Address. In the generated file, you will see two sheets (namely Employees & Addresses) will be generated.

Basic Excel header formatting is applied too.

## Models

```cs

public class Employee
{
    public int Id { get; set; }
    public string Name { get; set; }
    public DateTime DoB { get; set; }
    public decimal YearlySalary { get; set; }
    public Address Address { get; set; }
}

```

```cs

public class Address
{
    public int Id { get; set; }
    public string Line1 { get; set; }
    public string Line2 { get; set; }
    public string City { get; set; }
    public string State { get; set; }
    public string ZipCode { get; set; }
}

```

```cs

    public interface IEmployeeRepository
    {
        IEnumerable<Employee> GetEmployees();
    }

```

```cs

    public interface IExcelFileGenerator
    {
        void GenerateExcelFile(string fileName);
    }

```

## Project structure

![Project structure](http://oi64.tinypic.com/es6fsl.jpg)

## Client Program

```cs
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
```

## Screenshot of sample generated Excel file

![Screenshot of sample generated Excel file.](http://oi66.tinypic.com/2wcqpna.jpg)

For the whole source code, you clone and play with it. Since this solution uses [Office Open XML File Formats](https://www.ecma-international.org/publications/standards/Ecma-376.htm), it won't require you to install any driver or work with the *crazy Interop*.

*Note*: This isn't production ready code, so please customize to your specific need and do the required testing for correctness. This is only a demo project but can be easly used as starting point to create a real application.

Also feel free to use it anyway you want.
