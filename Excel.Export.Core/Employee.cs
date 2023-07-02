using System;

namespace Excel.Export.Core
{
    public class Employee
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime DoB { get; set; }
        public decimal YearlySalary { get; set; }
        public Address Address { get; set; }
    }
}
