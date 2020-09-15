using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ReadExcel
{
    public static class ExcelFile
    {
		public static string DownloadExcel(string path, string pathToDownload = null)
		{
			using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
				{
					string preffix = string.Empty;
					string name = fs.Name.Split('/').Last().Split('.').First();
					string ext = fs.Name.Split('.').Last();

					pathToDownload = string.IsNullOrEmpty(pathToDownload) ? "." : pathToDownload;

					if (!Directory.Exists(pathToDownload))
						Directory.CreateDirectory(pathToDownload);

					if (File.Exists(path))
						preffix = "_" + DateTime.Now.ToString("ddMMyyyyhhmmssmm");

					pathToDownload = $"{pathToDownload}/{name}{preffix}.{ext}";

					doc.SaveAs(pathToDownload);

					return pathToDownload;
				}
			}
		}

		public static void ReadExcel(string path, int ignoreRows = 0)
		{
			using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    Worksheet sheet = worksheetPart.Worksheet;

                    var rows = sheet.Descendants<Row>();

					if(ignoreRows > 0)
						rows = rows.Skip(ignoreRows);

                    foreach (Row row in rows)
                    {
                        var cells = row.Elements<Cell>();

                        foreach (Cell c in cells)
                        {
                            if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                Console.WriteLine("Shared string {0}: {1}", ssid, str);
                            }
                            else if (c.CellValue != null)
                            {
                                Console.WriteLine("Cell contents: {0}", c.CellValue.Text);
                            }
                        }

                        Console.WriteLine();
                    }
                }
                
            }
		}

         static void GenerateExcelConPanelesFijos()
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create("Ejemplo2.xlsx", SpreadsheetDocumentType.Workbook)) 
            {
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // rId must be unique within the spreadsheet. 
                // You might be able to use the SpreadSheetDocument.Parts.Count() to do this.
                // i.e. string relationshipID = "rId" + (spreadsheetDocument.Parts.Count() + 1).ToString();
                string rId = "rId6";

                // Sheet.Name and Sheet.SheetId must be unique within the spreadsheet.
                workbookPart.Workbook.Sheets = new Sheets();

                Sheet sheet = new Sheet() { Name = "Sheet4", SheetId = 4U, Id = rId };
                workbookPart.Workbook.Sheets.Append(sheet);

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>(rId);

                Worksheet worksheet = new Worksheet();

                // I don't know what SheetDimension.Reference is used for, it doesn't seem to change the resulting xml.
                SheetDimension sheetDimension = new SheetDimension() { Reference = "A1:A3" };
                SheetViews sheetViews = new SheetViews();
                // If more than one SheetView.TabSelected is set to true, it looks like Excel just picks the first one.
                SheetView sheetView = new SheetView() { TabSelected = false, WorkbookViewId = 0U };

                // I don't know what Selection.ActiveCell is used for, it doesn't seem to change the resulting xml.
                Selection selection = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };
                sheetView.Append(selection);
                sheetViews.Append(sheetView);

                SheetView sv = sheetViews.GetFirstChild<SheetView>();
                Selection sl = sv.GetFirstChild<Selection>();
                Pane pn = new Pane(){ VerticalSplit = 1D, TopLeftCell = "A2", ActivePane = PaneValues.BottomLeft, State = PaneStateValues.Frozen };
                sv.InsertBefore(pn,sl);
                sl.Pane = PaneValues.BottomLeft;

                SheetFormatProperties sheetFormatProperties = new SheetFormatProperties() { DefaultRowHeight = 15D };

                SheetData sheetData = new SheetData();

                // I don't know what the InnerText of Row.Spans is used for. It doesn't seem to change the resulting xml.
                Row row = new Row() { RowIndex = 1U, Spans = new ListValue<StringValue>() { InnerText = "1:3" } };

                Cell cell1 = new Cell() { CellReference = "A1", DataType = CellValues.Number, CellValue = new CellValue("99") };
                Cell cell2 = new Cell() { CellReference = "B1", DataType = CellValues.Number, CellValue = new CellValue("55") };
                Cell cell3 = new Cell() { CellReference = "C1", DataType = CellValues.Number, CellValue = new CellValue("33") };

                row.Append(cell1);
                row.Append(cell2);
                row.Append(cell3);

                sheetData.Append(row);
                PageMargins pageMargins = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.7D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };

                worksheet.Append(sheetDimension);
                worksheet.Append(sheetViews);
                worksheet.Append(sheetFormatProperties);
                worksheet.Append(sheetData);
                worksheet.Append(pageMargins);

                worksheetPart.Worksheet = worksheet;
            }
        }
	}
}