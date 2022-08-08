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
        Task<FileDownloadDto> GenerateExcelCustomize(
           DataTable dataTable,
           List<KeyValuePair<string, string>> displayFields,
           string fileName = "ResultFile",
           int[] columnWidthArray = null);
    }
}
