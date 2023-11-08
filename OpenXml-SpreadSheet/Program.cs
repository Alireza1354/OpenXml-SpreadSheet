using ClosedXML.Excel;
using OpenXml_SpreadSheet;


string filePath = @"C:\Users\karim\Documents\SampleData.xlsx";
string sheetName = "sheet2";

using (XLWorkbook workbook = new XLWorkbook(filePath))
{
    if (workbook.TryGetWorksheet(sheetName, out IXLWorksheet xLWorksheet))
    {
        //Cell(Row, Column)
        IXLColumn firstColumnUsed = xLWorksheet.FirstColumnUsed();
        IXLRow firstRowUsed = xLWorksheet.FirstRowUsed();
        IXLRows rowUsed = xLWorksheet.RowsUsed();
        var conutRowUsed = xLWorksheet.RowsUsed().Count();

        var firstColumNumber = firstColumnUsed.RangeAddress.FirstAddress.ColumnNumber;
        var firstRowNumber = firstRowUsed.RangeAddress.FirstAddress.RowNumber;
        Console.WriteLine($"firstColumn: {firstColumNumber} --- firstRowNumber {firstRowNumber}");


        var firstCellUsedValue = workbook.Worksheet(sheetName).Cell(firstRowNumber, firstColumNumber).Value;

        Dictionary<string, string> apiResult = await Api.GetData(firstCellUsedValue.ToString());

        var i = firstRowNumber - 1;
        var j = firstColumNumber;
        foreach (KeyValuePair<string, string> keyValuePair in apiResult)
        {
            var key = keyValuePair.Key;
            var value = keyValuePair.Value;

            workbook.Worksheet(sheetName).Cell(row: i, column: ++j).Value = key;
            workbook.Worksheet(sheetName).Cell(row: i + 1, column: j).Value = value;

        }
        IXLColumns colUsed = xLWorksheet.ColumnsUsed();
        var countColumnUsed = xLWorksheet.ColumnsUsed().Count();

        for (global::System.Int32 r = 1; r < conutRowUsed; r++)
        {

            var otheValue = workbook.Worksheet(sheetName).Cell(firstRowNumber + r, firstColumNumber).Value;

            apiResult.Clear();
            apiResult = await Api.GetData(otheValue.ToString());

            foreach (KeyValuePair<string, string> keyValuePair in apiResult)
            {
                var value = keyValuePair.Value;

                workbook.Worksheet(sheetName).Cell(row: firstRowNumber + r, column: firstColumNumber + 1).Value = value;
                firstColumNumber += 1;
                Console.WriteLine($"i: {firstRowNumber + r} -- j: {firstColumNumber}");
            }

            firstColumNumber -= countColumnUsed - 1;

        }

        workbook.Save();

        Console.WriteLine($"***Finished******");

    }

}


