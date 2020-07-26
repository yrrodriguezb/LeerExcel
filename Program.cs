namespace ReadExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"Financial Sample.xlsx";

            fileName = ExcelFile.DownloadExcel(fileName);

            ExcelFile.ReadExcel(fileName, 700);
        }
    }
}
