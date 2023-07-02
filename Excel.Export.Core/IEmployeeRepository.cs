using System.Collections.Generic;

namespace Excel.Export.Core
{
    public interface IEmployeeRepository
    {
        IEnumerable<Employee> GetEmployees();
    }
}
