using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace ExcelUtilities
{
  public static class ExcelMethods
  {
    /// <summary>
    /// Standard Blue color defined as a standard for all the worksheets.
    /// </summary>
    private static readonly Color standardBlue = Color.FromArgb(0, 45, 128);

    /// <summary>
    /// Convert an Excel workbook to an array of bytes.
    /// </summary>
    /// <param name="workbook">The name of the workbook.</param>
    /// <returns>An Excel workbook as an array of byte.</returns>
    public static byte[] ConvertExcelToByteArray(Workbook workbook)
    {
      string dataCompressed = ImportHelper.CompressObject(workbook);
      byte[] result = ObjectCompressionHelper.CompressObject(dataCompressed);
      return result;
    }

    /// <summary>
    /// Decompress an array of bytes from a string to an array.
    /// </summary>
    /// <param name="arrayOfByte">The array of bytes to be decompressed as a string.</param>
    /// <returns>An array of bytes.</returns>
    public static byte[] UnzipByteArray(string arrayOfByte)
    {
      var decompressedArray = ImportHelper.DecompressObject<byte[]>(arrayOfByte);
      return decompressedArray;
    }

    /// <summary>
    /// Build a new Excel File and fill it with data with an Export Global List Format.
    /// </summary>
    /// <param name="data">The list of data to be filled.</param>
    /// <param name="sheetName">The name of the worksheet.</param>
    /// <returns>An array of byte with the Excel file.</returns>
    public static byte[] BuildExcelFile(List<string> data, string sheetName)
    {
      Workbook workbook = CreateWorkbook(sheetName);
      List<string> headers = new List<string>();
      headers.Add("Header1");
      headers.Add("Header2");
      headers.Add("Header3");
      headers.Add("Header4");
      headers.Add("Header5");
      headers.Add("Header6");
      headers.Add("Header7");


      AddHeader(workbook, headers);
      // todo to complete code
      return new byte[0]; // change accordingly
    }

    /// <summary>
    /// Add headers to a worksheet.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="headers">A list of strings with the header names.</param>
    /// <param name="worksheetPosition">The position of the worksheet to be used as an integer starting with zero.</param>
    /// <param name="startingColumn">The position of the column to input data as an integer starting with zero. The default value is zero.</param>
    /// <param name="startingRow">The position of the row to input data as an integer starting with zero. The default value is zero.</param>
    /// <returns></returns>
    public static Workbook AddHeader(Workbook workbook, List<string> headers, int worksheetPosition = 0, int startingColumn = 0, int startingRow = 0)
    {
      // Fill the headers of the worksheet
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      Cells cells = worksheet.Cells;
      int counter = 0;
      foreach (string item in headers)
      {
        cells[startingRow, startingColumn + counter].PutValue(item);
        counter++;
      }

      return workbook;
    }

    /// <summary>
    /// Set a style to cells in a worksheet in a workbook.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="color">The color of the cell. If color is null then the standard blue color is used.</param>
    /// <param name="worksheetPosition">The position of the worksheet. The default value is zero.</param>
    /// <param name="fontIsBold">Is the font in bold.</param>
    /// <param name="startingColumn">The starting column number as an integer. the default value is zero.</param>
    /// <param name="endingColum">The ending column number as an integer. the default value is zero.</param>
    /// <param name="startingRow">The starting row number as an integer. the default value is zero.</param>
    /// <returns>A workbook with style applied.</returns>
    public static Workbook SetStyle(Workbook workbook, Color? color, int worksheetPosition = 0, bool fontIsBold = false, int startingColumn = 0, int endingColum = 0, int startingRow = 0)
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      Cells cells = worksheet.Cells;
      Cell plage = worksheet.Cells["A1"];
      Style headerStyle = plage.GetStyle();
      if (color == null)
      {
        headerStyle.ForegroundColor = standardBlue;
      }
      else
      {
        headerStyle.ForegroundColor = (Color)color;
      }

      headerStyle.Pattern = BackgroundType.Solid;
      headerStyle.Font.IsBold = fontIsBold;

      for (int i = startingColumn; i <= endingColum; i++)
      {
        cells[startingRow, startingColumn + i].SetStyle(headerStyle);
      }

      return workbook;
    }

    public static Workbook ImportData(Workbook workbook)
    {
      // todo
      return workbook;
    }

    /// <summary>
    /// Set to auto fit columns in a worksheet in a workbook.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetPosition">The position of the worksheet. The default value is zero.</param>
    /// <returns>A workbook with auto fit column applied in a worksheet.</returns>
    public static Workbook AutoFitColumns(Workbook workbook, int worksheetPosition = 0)
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      worksheet.AutoFitColumns();
      return workbook;
    }

    /// <summary>
    /// Exports the data and returns the byte array containing the result.
    /// </summary>
    /// <returns>The byte array containing the exported data.</returns>
    public static byte[] Export(List<ExportFormat> data, string sheetName)
    {
      byte[] excelBytes;

      using (Workbook workbook = new Workbook())
      {
        Worksheet worksheet = null;

        if (workbook.Worksheets.Count == 0)
        {
          worksheet = workbook.Worksheets.Add(sheetName);
        }
        else
        {
          worksheet = workbook.Worksheets[0];

          worksheet.Name = sheetName;
        }

        FeedDataToExcel(worksheet, data);
        // save excel to test format
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;
        workbook.FileFormat = FileFormatType.Excel97To2003;
        string excelFileName = $"{sheetName}.xlsx";
        workbook.Save(excelFileName, SaveFormat.Xlsx);
        excelBytes = ConvertExcelToByteArray(workbook);
        //excelBytes = workbook.SaveToStream().ToArray();
        worksheet.Dispose();
      }

      return excelBytes;
    }

    /// <summary>
    /// Prepares the worksheet to convert from the list of data to export.
    /// </summary>
    /// <param name="data">The list containing the data.</param>
    /// <param name="worksheet">The worksheet to fill with the data informations.</param>
    public static void FeedDataToExcel(Worksheet worksheet, List<ExportFormat> data)
    {
      try
      {
        // Fill the header of the file
        Cells cells = worksheet.Cells;
        cells[0, 0].PutValue("Header1");
        cells[0, 1].PutValue("Header2");
        cells[0, 2].PutValue("Header3");
        cells[0, 3].PutValue("Header4");
        cells[0, 4].PutValue("Header5");
        cells[0, 5].PutValue("Header6");
        cells[0, 6].PutValue("Header7");
        Cell plage = worksheet.Cells["A1"];
        Style headerStyle = plage.GetStyle();
        headerStyle.ForegroundColor = standardBlue;
        headerStyle.Pattern = BackgroundType.Solid;
        headerStyle.Font.IsBold = true;

        for (int i = 0; i < 8; i++)
        {
          cells[0, i].SetStyle(headerStyle);
        }

        cells.ImportCustomObjects(
            data.OrderBy(f => f.Header1).ToList(),
            new string[] { "Header1", "Header2", "Header3", "Header4", "Header5", "Header6", "Header7" },
            false,
            1,
            0,
            data.Count,
            true,
            "dd/mm/yyyy",
            false);
        worksheet.AutoFitColumns();
      }
      catch (Exception exception)
      {
        throw new Exception("Error while parsing data to export", exception);
      }
    }

    public static Worksheet CreateWorksheet(string worksheetName)
    {
      Workbook workbook = new Workbook();
      Worksheet worksheet = null;

      if (workbook.Worksheets.Count == 0)
      {
        worksheet = workbook.Worksheets.Add(worksheetName);
      }
      else
      {
        worksheet = workbook.Worksheets[0];

        worksheet.Name = worksheetName;
      }

      return worksheet;
    }

    /// <summary>
    /// Create a new workbook.
    /// </summary>
    /// <param name="worksheetName">The name of the first worksheet. By default, it is sheet1.</param>
    /// <returns></returns>
    public static Workbook CreateWorkbook(string worksheetName = "sheet1")
    {
      Workbook workbook = new Workbook();
      Worksheet worksheet = null;

      if (workbook.Worksheets.Count == 0)
      {
        worksheet = workbook.Worksheets.Add(worksheetName);
      }
      else
      {
        worksheet = workbook.Worksheets[0];

        worksheet.Name = worksheetName;
      }

      return workbook;
    }

    /// <summary>
    /// Add a worksheet to an existing workbook.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetName">The name of the worksheet to be added.</param>
    /// <returns>The workbook with and added worksheet.</returns>
    public static Workbook AddWorksheet(Workbook workbook, string worksheetName = "sheet2")
    {
      Worksheet worksheet = null;
      worksheet = workbook.Worksheets.Add(worksheetName);
      worksheet.Name = worksheetName;
      return workbook;
    }
  }
}
