using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using GeneradorDeFirmaInduban.Models;
using GeneradorDeFirmaInduban.Services;
using LinqToExcel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace GeneradorDeFirmaInduban.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class GenerateSigningController : ControllerBase
    {
        private readonly ILogger<GenerateSigningController> _logger;
        private readonly IWebHostEnvironment _env;
        private readonly IEmployeeInfoService _employeeInfoService;
        
        public GenerateSigningController(ILogger<GenerateSigningController> logger, IWebHostEnvironment env, IEmployeeInfoService employeeInfoService)
        {
            _logger = logger;
            _env = env;
            _employeeInfoService = employeeInfoService;
        }
  
        [HttpGet]
        [Route("generate")]
        public IActionResult Generate(string employeeCode)
        {

            EmployeeInfo info = _employeeInfoService.GetUserInfo(employeeCode);

            if (info != null)
            {
                _employeeInfoService.GenerateSignature(info);
            }
            else
            {
                return Ok(null);
            }

            return Ok(info);
        }
        [HttpGet]
        [Route("downloadfile")]
        public ActionResult DownloadSignatureFile(string employeeCode)
        {
            try
            {
                string inputPath = _env.ContentRootPath + @"\files\generated\firma"+ employeeCode + ".docx";

                var file = Path.Combine(inputPath);

                byte[] fileBytes = System.IO.File.ReadAllBytes(file);


               return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "firma"+ employeeCode + ".docx");
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
    }
}
