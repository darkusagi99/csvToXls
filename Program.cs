using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace csvToXls
{
    class Program
    {
        static void Main(string[] args) { 

        if(args.Count() < 1) {
            Console.WriteLine("Please enter a file path:");
            args = new string[1];
            args[0] = Console.ReadLine();
        }
        var path = args[0];

        if (!File.Exists(path) || !path.ToUpper().Contains("CSV")) {
        Console.WriteLine("Must provide a valid CSV");
        System.Threading.Thread.Sleep(500);
        return;
        }

        List<dynamic> issues;


        DataTable csvDataTable = ConvertCSVtoDataTable(path);

        //using (var wb = new ClosedXML.Excel.XLWorkbook()) {
        //    wb.AddWorksheet(table, "Sheet1");
        //    foreach (var ws in wb.Worksheets)
        //    {
        //    ws.Columns().AdjustToContents();
        //    }
            var outputPath = path.Substring(0, path.Length - 3) + "xlsx";
            //    wb.SaveAs(output);

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(outputPath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Create sheetdatas
            SheetData xlSheetData = new SheetData();

            Row xlHeaderRow = new Row();
            foreach (DataColumn col in csvDataTable.Columns)
            {
                object cellData = col.ColumnName;
                Cell xlCell = null;
                if (cellData != null)
                {
                    xlCell = new Cell(new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(cellData.ToString()))) { DataType = CellValues.InlineString };
                }
                else
                {
                    xlCell = new Cell(new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(String.Empty))) { DataType = CellValues.InlineString };
                }
                xlHeaderRow.Append(xlCell);
            }
            xlSheetData.Append(xlHeaderRow);
            
            // Add content as lines
            foreach (DataRow row in csvDataTable.Rows)
            {
                Row xlRow = new Row();
                foreach (DataColumn col in csvDataTable.Columns)
                {
                    object cellData = row[col];
                    Cell xlCell = null;
                    if (cellData != null)
                    {
                        xlCell = new Cell(new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(cellData.ToString()))) { DataType = CellValues.InlineString };
                    }
                    else
                    {
                        xlCell = new Cell(new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(String.Empty))) { DataType = CellValues.InlineString };
                    }
                    xlRow.Append(xlCell);
                }
                xlSheetData.Append(xlRow);
            }

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(xlSheetData);

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "CSV"
            };
            sheets.Append(sheet);

            



            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();

            Console.WriteLine("wrote to : " + outputPath);
            System.Threading.Thread.Sleep(500);
            return;
        //}
    }

    public static DataTable ConvertCSVtoDataTable(string strFilePath)
    {

        DataTable dt = new DataTable();
        StreamReader sr = new StreamReader(strFilePath);
        string[] headers = sr.ReadLine().Split(';');
        foreach (string header in headers)
        {
            dt.Columns.Add(header);
        }
        while (!sr.EndOfStream)
        {
            string[] rows = System.Text.RegularExpressions.Regex.Split(sr.ReadLine(), ";(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
            DataRow dr = dt.NewRow();
            for (int i = 0; i < headers.Length; i++)
            {
                dr[i] = rows[i];
            }
            dt.Rows.Add(dr);
        }
        return dt;
    }

    }


}

