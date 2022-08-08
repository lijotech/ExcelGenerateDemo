using ExcelGenerate.DTO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGenerate.Services
{
    public class ExcelGenerateService:IExcelGenerateService
    {
        public Task<FileDownloadDto> GenerateExcel(DataTable dataTable)
        {
            return Task.FromResult( new FileDownloadDto { });
        }
    }
}
