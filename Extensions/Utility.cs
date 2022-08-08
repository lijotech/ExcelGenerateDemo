using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGenerate.Extensions
{
    public static class Utility
    {
        public static DataTable GenerateDatatableWithData(int columsNumber, int rowsCount)
        {

            DataTable table = new DataTable();
            for (int d = 0; d <= columsNumber; d++)
            {
                table.Columns.Add("Column" + d, typeof(string));
            }

            for (int r = 0; r <= rowsCount; r++)
            {
                DataRow dataRow = table.NewRow();
                for (int d = 0; d <= columsNumber; d++)
                {
                    dataRow[d] = "cell" + d;
                }
                table.Rows.Add(dataRow);
            }

            return table;
        }
    }
}
