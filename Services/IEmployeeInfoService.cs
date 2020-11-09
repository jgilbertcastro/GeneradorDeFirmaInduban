using GeneradorDeFirmaInduban.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GeneradorDeFirmaInduban.Services
{
    public interface IEmployeeInfoService
    {
        public EmployeeInfo GetUserInfo(string employeeCode);
        public void GenerateSignature(EmployeeInfo employeeInfo);
        public void SaveDateGenerated(string employeeCode, DateTime date);
    }
}
