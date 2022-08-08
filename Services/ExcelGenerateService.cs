using ExcelGenerate.DTO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGenerate.Services
{
    public class ExcelGenerateService : IExcelGenerateService
    {
        public static readonly string LOGOIMGPATH = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "assets", "logo.jpg");
        public Task<FileDownloadDto> GenerateExcel(DataTable dataTable)
        {
            return Task.FromResult(new FileDownloadDto { });
        }

        public Task<FileDownloadDto> GenerateExcelCustomize(
           DataTable dataTable,
           List<KeyValuePair<string, string>> displayFields,
           string fileName = "ResultFile",
           int[] columnWidthArray = null)
        {
            FileDownloadDto resultFile = new FileDownloadDto
            {
                MimeType = "application/vnd.ms-excel"
            };
            using (var pck = new ExcelPackage())
            {
                DataTable dt = dataTable;
                if (columnWidthArray == null || !columnWidthArray.Any())
                    columnWidthArray = Enumerable.Repeat(20, dataTable.Columns.Count + 1).ToArray();



                var wsEnum = pck.Workbook.Worksheets.Add(fileName);
                int firstColumnNumber = 3;
                for (int i = 0; i < (dt.Columns.Count); i++)
                {
                    wsEnum.Column(i + (firstColumnNumber)).Width = columnWidthArray[i];
                }
                wsEnum.View.ShowGridLines = false;

                char beginColumnAlphabet = 'C';
                char endColumnAlphabet = incrementCharacter(beginColumnAlphabet, dt.Columns.Count - 1);
                char mergeColumnAlphabet = incrementCharacter(beginColumnAlphabet, dt.Columns.Count - 2);


                //styling the header row               
                wsEnum.Cells[$"{beginColumnAlphabet}7:{endColumnAlphabet}7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                wsEnum.Cells[$"{beginColumnAlphabet}7:{endColumnAlphabet}7"].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                wsEnum.Cells[$"{beginColumnAlphabet}7:{endColumnAlphabet}7"].Style.Font.Bold = true;
                wsEnum.Cells[$"{beginColumnAlphabet}7:{endColumnAlphabet}7"].Style.Font.Color.SetColor(Color.Black);

                //styling of the title and mergeing the cells
                wsEnum.Cells[$"{beginColumnAlphabet}5:{mergeColumnAlphabet}6"].Merge = true;
                wsEnum.Cells[$"{beginColumnAlphabet}5"].Style.Font.Bold = true;
                wsEnum.Cells[$"{beginColumnAlphabet}5"].Style.Font.Size = 20;
                wsEnum.Cells[$"{beginColumnAlphabet}5"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                wsEnum.Cells[$"{beginColumnAlphabet}5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                foreach (var item in displayFields)
                {
                    switch (item.Key)
                    {
                        case "TITLE_FIELD":
                            wsEnum.Cells[$"{beginColumnAlphabet}5"].Value = item.Value;
                            break;
                        case "DATE_FIELD":
                            wsEnum.Cells[$"{endColumnAlphabet}4"].Value = item.Value;
                            break;
                        case "TIME_FIELD":
                            wsEnum.Cells[$"{endColumnAlphabet}5"].Value = item.Value;
                            break;
                        case "CREATEDBY_FIELD":
                            wsEnum.Cells[$"{endColumnAlphabet}6"].Value = item.Value;
                            break;
                    }

                }

                //setting logo from assets folder
                var picture = wsEnum.Drawings.AddPicture("logo", LOGOIMGPATH);
                picture.SetPosition(1, 0, 2, 0);
                picture.SetSize(150, 95);

                //loading data from datatable
                wsEnum.Cells[$"{beginColumnAlphabet}7"].LoadFromDataTable(dt, true, TableStyles.None);

                char columnPosition = beginColumnAlphabet;
                foreach (DataColumn dataColumn in dt.Columns)
                {
                    wsEnum.Cells[$"{columnPosition}7"].Value = (dataColumn.ColumnName)
                        .Replace("[", "").Replace("]", "");
                    columnPosition = incrementCharacter(columnPosition);
                }

                // Assign borders from the starting cell to end of rows
                var modelRows = (dt.Rows.Count + 1) + 6;
                string modelRange = $"{beginColumnAlphabet}7:{endColumnAlphabet }" + modelRows.ToString();
                var modelTable = wsEnum.Cells[modelRange];
                modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                modelTable.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                //Can be used in case we want auto fit size of column based on data
                //wsEnum.Cells[wsEnum.Dimension.Address].AutoFitColumns();
                resultFile.Attachment = pck.GetAsByteArray();
                resultFile.FileName = string.Format("{0}_{1}.xlsx", fileName, System.DateTime.Now.ToString("yyyyMMddHHmmssffff"));
            }
            return Task.FromResult(resultFile);
        }

        char incrementCharacter(char input, int increment = 1)
        {
            return (input == 'z' ? 'a' : (char)(input + increment));
        }


    }
}
