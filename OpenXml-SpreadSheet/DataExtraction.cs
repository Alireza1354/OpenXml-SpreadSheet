using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML;
using ClosedXML.Excel;
using System.Data;
using System.Linq.Expressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXml_SpreadSheet
{
    internal class DataExtraction
    {

        public static async Task<DataRow> GetRegionRow(string RegionId)
        {
            var regions = new List<string>();
            using var wb =
                new XLWorkbook($"C:\\Users\\karim\\Documents\\TbRegions.xlsx");

            var ws = wb.Worksheet("sheet1");

            var firstRowUsed = ws.FirstRowUsed();
            var regionRow = firstRowUsed.RowUsed();

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
            foreach (var r in regionTable.DataRange.Rows())
            {
                dtRow = dt.NewRow();
                dtRow["RegionID"] = r.Field("RegionID").GetString();
                dtRow["CenterID"] = r.Field("CenterID").GetString();
                dtRow["CenterName"] = r.Field("CenterName").GetString();
                dtRow["RegionName"] = r.Field("RegionName").GetString();
                dt.Rows.Add(dtRow);
            }

            string expression = "RegionID" + "=" + RegionId;
            try
            {
                DataRow row = dt.Select(expression).FirstOrDefault();
                return row;
            }
            catch (Exception ex)
            {
                dtRow = dt.NewRow();
                dtRow["RegionID"] = "";
                dtRow["CenterID"] = "";
                dtRow["CenterName"] = "";
                dtRow["RegionName"] = "";
                Console.WriteLine(ex.Message);
                return dtRow;
            }


            
        }
    }
}