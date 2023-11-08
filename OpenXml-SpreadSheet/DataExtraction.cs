using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML;
using ClosedXML.Excel;
using System.Data;
using System.Linq.Expressions;

namespace OpenXml_SpreadSheet
{
    internal class DataExtraction
    {

        public static async Task GetDataTable()
        {
            var regions = new List<string>();
            using var wb =
                new XLWorkbook($"C:\\Users\\karim\\source\\repos\\Alireza1354\\OpenXml-SpreadSheet\\OpenXml-SpreadSheet\\Rgions.xlsx");
            var ws = wb.Worksheet("sheet1");
            var firstRowUsed = ws.FirstRowUsed();
            var regionRow = firstRowUsed.RowUsed();
            // Move to the next row (it now has the titles)
            // regionRow = regionRow.RowBelow();

            // First possible address of the Region table
            var firstPossibleAddress =
                ws.Row(regionRow.RowNumber())
                .FirstCell()
                .Address;

            // Last possibleAddress of the Region table
            var lastPossibleAddress = ws.LastCellUsed().Address;

            // Get a range with the remainder of the worksheet data (the range used)
            var regionRange = ws.Range(firstPossibleAddress, lastPossibleAddress);

            // Treat the range as a table (to be able to use the column names)
            IXLTable regionTable = regionRange.AsTable();

            //foreach ( var field in fields )
            //{
            //    field.Column
            //}
            // Get the list of company names
            //regions = regionTable.DataRange.Rows()
            //    .Select(r => r.Field("RegionName").GetString()).Where(x => x != null)
            //    .ToList();

            DataTable dt = new DataTable();
            DataColumn dataColumn = new DataColumn()
            {
                ColumnName = "Regionid",
                DataType = typeof(string),
                Unique = true,
            };
            dt.Columns.Add(dataColumn);

            dataColumn = new DataColumn()
            {
                ColumnName = "Centerid",
                DataType = typeof(string),
            };
            dt.Columns.Add(dataColumn);

            dataColumn = new DataColumn()
            {
                ColumnName = "CenterName",
                DataType = typeof(string),
            };
            dt.Columns.Add(dataColumn);

            dataColumn = new DataColumn()
            {
                ColumnName = "RegionName",
                DataType = typeof(string),
            };
            dt.Columns.Add(dataColumn);

            DataRow dtRow;
            foreach (var row in regionTable.DataRange.Rows())
            {
                dtRow = dt.NewRow();
                dtRow["Regionid"] = row.Field("Regionid").GetString();
                dtRow["Centerid"] = row.Field("Centerid").GetString();
                dtRow["CenterName"] = row.Field("CenterName").GetString();
                dtRow["RegionName"] = row.Field("RegionName").GetString();
                dt.Rows.Add(dtRow);
            }


            string expression = "Regionid = 1801";
            DataRow[] foundRows;
            foundRows = dt.Select(expression);
            var reg = foundRows[0]["RegionName"].ToString();
            var cap = regions.Capacity;
        }
    }
}
