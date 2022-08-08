using ExcelGenerate.DTO;
using ExcelGenerate.Extensions;
using ExcelGenerate.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGenerate.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelGenerateController : ControllerBase
    {
        private readonly IExcelGenerateService _excelGenerateService;
        public ExcelGenerateController(IExcelGenerateService excelGenerateService)
        {
            _excelGenerateService = excelGenerateService;
        }

        [HttpGet("GenerateExcel")]
        public async Task<IActionResult> GenerateExcel()
        {
            try
            {
                var result = await _excelGenerateService.GenerateExcel(Utility.GenerateDatatableWithData(4, 3));

                return File(result.Attachment, result.MimeType, result.FileName);
            }

            catch (Exception ex)
            {

                var response = new
                {
                    Msg = "Processing request failed.",
                    Errorlst = new List<ErrorMessage>() {
                        new ErrorMessage() { Error = ex.Message } }
                };
                return StatusCode(StatusCodes.Status500InternalServerError, response);
            }
        }
    }
}
