using GeneradorDeFirmaInduban.Models;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using TemplateEngine.Docx;

namespace GeneradorDeFirmaInduban.Services
{
    public class EmployeeInfoService : IEmployeeInfoService
    {
        private readonly IWebHostEnvironment _env;

        public EmployeeInfoService(IWebHostEnvironment env)
        {
            _env = env;
        }
        public void GenerateSignature(EmployeeInfo employeeInfo)
        {
            string inputPath = _env.ContentRootPath + @"\files\templates\firma aprobada.docx";
            string outputPath = _env.ContentRootPath + @"\files\generated\firma" + employeeInfo.Id+".docx";
            File.Delete(outputPath);
            File.Copy(inputPath, outputPath);

            var valuesToFill = new Content(
                new FieldContent("nombre", employeeInfo.Nombre),
                new FieldContent("apellido",employeeInfo.Apellido),
                new FieldContent("posicion",employeeInfo.Posicion),
                new FieldContent("ext",employeeInfo.Ext),
                new FieldContent("flota",employeeInfo.Flota));

            using (var outputDocument = new TemplateProcessor(outputPath)
                .SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            }

            SaveDateGenerated(employeeInfo.Id, DateTime.Now);
        }

        public EmployeeInfo GetUserInfo(string employeeCode)
        {
            EmployeeInfo info = null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(_env.ContentRootPath + @"\files\templates\empleados.xlsx")))
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets[0];

                int rowStart = workSheet.Dimension.Start.Row;
                int rowEnd = workSheet.Dimension.End.Row;

                string cellRange = rowStart.ToString() + ":" + rowEnd.ToString();

                var searchCell = from cell in workSheet.Cells[cellRange]
                                 where cell.Value.ToString() == employeeCode
                                 select cell.Start.Row;

                if (searchCell.Any())
                {
                    info = new EmployeeInfo();
                    int rowNum = searchCell.First();

                    info.Id = workSheet.GetValue(rowNum, 1).ToString();
                    info.Nombre = workSheet.GetValue(rowNum, 2).ToString();
                    info.Apellido = workSheet.GetValue(rowNum, 3).ToString();
                    info.Ext = workSheet.GetValue(rowNum, 4).ToString();
                    info.Flota = workSheet.GetValue(rowNum, 5).ToString();
                    info.Posicion = workSheet.GetValue(rowNum, 6).ToString();
                }
            }

            return info;
        }

        public void SaveDateGenerated(string employeeCode, DateTime date)
        {
            string text = employeeCode + " " + date.ToString();
            try
            {
                string filePath = _env.ContentRootPath + @"\files\logs\savesignaturedategenerated.txt";

                using (StreamWriter write = new StreamWriter(filePath, true))
                {
                    write.WriteLine(text);
                    write.Flush();
                    write.Close();
                }
            }
            catch
            {
                
            }
        }
    }
}
