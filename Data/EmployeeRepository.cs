using Excel.Export.Core;
using System;
using System.Collections.Generic;

namespace Data
{
    public class EmployeeRepository : IEmployeeRepository
    {
        private const int TOTAL_EMPLOYEES_COUNT = 10;
        private static IList<Employee> _employees;

        /// <summary>
        /// Gets employess
        /// </summary>
        /// <returns>List of employees</returns>
        public IEnumerable<Employee> GetEmployees()
        {
            //FIXME  -- for real application - you should query database and get your record. For demo however, that will be an overhead and hence in memory object collection should be fine.
            if (_employees == null)
            {
                _employees = new List<Employee>();
                Random random = new Random();
                for (int i = 0; i < TOTAL_EMPLOYEES_COUNT; i++)
                {
                    _employees.Add(new Employee
                    {
                        Id = i + 1,
                        Name = $"{i}-Employee",
                        DoB = new DateTime(1999, 1, 1).AddMonths(i + 1),
                        YearlySalary = random.Next(100000, 350000),
                        Address = new Address
                        {
                            Id = i + 1,
                            Line1 = Guid.NewGuid().ToString().Substring(0, 20).Replace('-', ' '),
                            Line2 = i % 3 == 0 ? $"Suite {Guid.NewGuid().ToString().Substring(0, 3)}" : $"Apt# {Guid.NewGuid().ToString().Substring(0, 2)}",
                            City = $"Random city {i}",
                            State = $"{i} State",
                            ZipCode = random.Next(10000, 99999).ToString()
                        }
                    });
                }
            }

            return _employees;
        }

    }
}
