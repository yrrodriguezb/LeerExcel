using System;
using System.IO;
using System.Linq;
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
	}
}