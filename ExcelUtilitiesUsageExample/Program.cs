using ExcelUtilities;
using System;
using System.Linq;

namespace ExcelUtilitiesUsageExample
{
  internal class Program
  {
    static void Main()
    {
      Action<string> Display = Console.WriteLine;
      Display("Example of usage of Excel Utilities library");
      var workbook = ExcelMethods.CreateWorkbook();
      string sheetName = "example";

      if (!workbook.Worksheets.Names.ToList().Any( n => n.Text == sheetName))
      {
        workbook.Worksheets.Add(sheetName);
      }

      var worksheet = workbook.Worksheets[0];
      worksheet.Cells["A1"].Value = "test1";
      worksheet.Hyperlinks.Add("A2", 1, 1, "http://www.aspose.com");
      string fileName = "test.xlsx";
      workbook.Save(fileName);
      ProcessHelper.ExecuteProcess(fileName);
      Display("Press any key to exit:");
      Console.ReadLine();
    }
  }
}
