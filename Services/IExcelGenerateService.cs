using ExcelGenerate.DTO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGenerate.Services
{
    public interface IExcelGenerateService
    {

        Task<FileDownloadDto> GenerateExcel(DataTable dataTable);
    }
}
