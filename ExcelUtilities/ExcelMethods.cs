﻿using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Font = Aspose.Cells.Font;

namespace ExcelUtilities
{
  public static class ExcelMethods
  {
    /// <summary>
    /// Standard Blue color defined as a standard for all the worksheets.
    /// </summary>
    private static readonly Color standardBlue = Color.FromArgb(0, 45, 128);
    private static readonly Color? defaultBlue;

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

    /// <summary>
    /// Get a style from a cell.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetPosition">the position of the worksheet. The default value is zero.</param>
    /// <param name="rowNumber">The row number as an integer. The default value is zero.</param>
    /// <param name="columnNumber">The column number as an integer. The default value is zero.</param>
    /// <returns>the style of the cell designated by rowNumber and columnNumber</returns>
    public static Style GetStyle(Workbook workbook, int worksheetPosition = 0, int rowNumber = 0, int columnNumber = 0)
    {
      //Create a Style object to fetch the Style of a Cell.
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      Style style = worksheet.Cells[rowNumber, columnNumber].GetStyle();
      return style;
    }

    /// <summary>
    /// Set the font, the size and the color of a range of cells.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetPosition">The position of the worksheet. The default value is zero.</param>
    /// <param name="styleRowNumber">The row number to get the style from. The default value is zero.</param>
    /// <param name="styleColumnNumber">The column number to get the style from. The default value is zero.</param>
    /// <param name="startingRow">The starting row number to apply the style. The default value is zero.</param>
    /// <param name="endingRow">The ending row number to apply the style. The default value is zero.</param>
    /// <param name="startingColumn">The starting column number to apply the style. The default value is zero.</param>
    /// <param name="endingColumn">The ending column number to apply the style. The default value is zero.</param>
    /// <param name="size">The size of the font. The default value is 11.</param>
    /// <param name="color">The color of the font. The default value is Black.</param>
    /// <param name="isBold">Is it in bold? The default value is false.</param>
    /// <param name="isItalic">Is it in italic? The default value is false.</param>
    /// <param name="fontName">The name of the font e.g. "Calibri" or "Times New Roman". The default value is Calibri.</param>
    /// <returns>A workbook with the requested style applied.</returns>
    public static Workbook SetFontSizeAndColor(Workbook workbook, int worksheetPosition = 0, int styleRowNumber = 0, int styleColumnNumber = 0, int startingRow = 0, int endingRow = 0, int startingColumn = 0, int endingColumn = 0, int size = 11, string color = "Black", bool isBold = false, bool isItalic = false, string fontName = "Calibri")
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      Cells cells = worksheet.Cells;

      //Create a Style object to fetch the Style of a Cell.
      Style style = worksheet.Cells[styleRowNumber, styleColumnNumber].GetStyle();

      //Create a Font object
      Font font = style.Font;

      //Set the name.
      font.Name = fontName;// examples: "Calibri" or "Times New Roman"

      //Set the font size.
      font.Size = size;

      //Set the font color
      if (color == "Black")
      {
        font.Color = Color.Black;
      }
      else
      {
        font.Color = Color.Black; // add other colors if needed
      }

      font.IsBold = isBold;
      font.IsItalic = isItalic;
      style.ForegroundColor = Color.White;
      style.Pattern = BackgroundType.Solid;

      for (int i = startingRow; i <= endingRow; i++)
      {
        for (int j = startingColumn; j <= endingColumn; j++)
        {
          cells[styleRowNumber + i, styleColumnNumber + j].SetStyle(style);
        }
      }

      return workbook;
    }

    /// <summary>
    /// Center the text of one or several cells.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="styleRowNumber">The number of the row to get the style.</param>
    /// <param name="styleColumnNumber">The number of the column to get the style.</param>
    /// <param name="worksheetPosition">The position of the worksheet starting with zero.</param>
    /// <param name="startingColumn">The column number to start from.</param>
    /// <param name="endingColumn">The column number to end.</param>
    /// <param name="startingRow">The row number to start from.</param>
    /// <param name="endingRow">The row number to end.</param>
    /// <param name="textAlignmentType">The type of the text alignment like Left, Right, Center, etc.</param>
    /// <returns>A workbook with text aligned.</returns>
    public static Workbook CenterColumn(Workbook workbook, int styleRowNumber = 0, int styleColumnNumber = 0, int worksheetPosition = 0, int startingColumn = 0, int endingColumn = 0, int startingRow = 0, int endingRow = 0, TextAlignmentType textAlignmentType = TextAlignmentType.Left)
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      Cells cells = worksheet.Cells;

      //Create a Style object to fetch the Style of a Cell.
      Style style = worksheet.Cells[styleRowNumber, styleColumnNumber].GetStyle();
      style.HorizontalAlignment = textAlignmentType;

      for (int i = startingRow; i <= endingRow; i++)
      {
        for (int j = startingColumn; j <= endingColumn; j++)
        {
          cells[i, j].SetStyle(style);
        }
      }

      return workbook;
    }

    /// <summary>
    /// Replace text on a range of cells.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetPosition">The position of the worksheet to be used.</param>
    /// <param name="startingColumn">The column number to start from.</param>
    /// <param name="endingColumn">The ending column number.</param>
    /// <param name="startingRow">The row number to start from.</param>
    /// <param name="endingRow">The ending row number.</param>
    /// <param name="oldText">The text to be replaced.</param>
    /// <param name="newText">The new text to replace the old text.</param>
    /// <returns>A workbook with text replaced.</returns>
    private static Workbook ReplaceText(Workbook workbook, int worksheetPosition, int startingColumn = 0, int endingColumn = 0, int startingRow = 0, int endingRow = 0, string oldText = "", string newText = "")
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      Cells cells = worksheet.Cells;
      for (int i = startingRow; i <= endingRow; i++)
      {
        for (int j = startingColumn; j <= endingColumn; j++)
        {
          if (cells[i, j].Value != null && cells[i, j].Value.ToString().ToUpper() == oldText.ToUpper())
          {
            cells[i, j].PutValue(newText);
          }
        }
      }

      return workbook;
    }

    /// <summary>
    /// Insert one or several rows.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetPosition">The position of the worksheet to be used.</param>
    /// <param name="rowNumber">The row number where a row will be inserted.</param>
    /// <param name="numberOfRowtoBeAdded">The number of row to be added. The default value is one.</param>
    /// <returns>A workbook with row added.</returns>
    private static Workbook InsertRow(Workbook workbook, int worksheetPosition, int rowNumber = 0, int numberOfRowtoBeAdded = 1)
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      worksheet.Cells.InsertRows(rowNumber, numberOfRowtoBeAdded);
      return workbook;
    }

    /// <summary>
    /// Copy the format of one cell to another one or a range of cells.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetPosition">The position of the worksheet to be used.</param>
    /// <param name="cellRowtoCopyFrom">The row number of the cell to copy the format from.</param>
    /// <param name="cellColumntoCopyFrom">The column number of the cell to copy the format from.</param>
    /// <param name="startingColumn">The column number to start from.</param>
    /// <param name="endingColumn">The ending column number.</param>
    /// <param name="startingRow">The row number to start from.</param>
    /// <param name="endingRow">The ending row number.</param>
    /// <returns>A workbook with a new format applied on one cell or several cells.</returns>
    public static Workbook CopyCellFormat(Workbook workbook, int worksheetPosition, int cellColumntoCopyFrom, int cellRowtoCopyFrom, int startingColumn = 0, int endingColumn = 0, int startingRow = 0, int endingRow = 0)
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      Cells cells = worksheet.Cells;

      //Create a Style object to fetch the Style of a Cell.
      Style style = worksheet.Cells[cellRowtoCopyFrom, cellColumntoCopyFrom].GetStyle();

      for (int i = startingRow; i <= endingRow; i++)
      {
        for (int j = startingColumn; j <= endingColumn; j++)
        {
          cells[i, j].SetStyle(style);
        }
      }

      return workbook;
    }

    /// <summary>
    /// Write text to one cell.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetPosition">The position of the worksheet to be used.</param>
    /// <param name="text">The text to be inserted.</param>
    /// <param name="cellRow">The row number of the cell to be written.</param>
    /// <param name="cellColumn">The column number of the cell to be written.</param>
    /// <returns>A workbook with the text insreted in a cell.</returns>
    public static Workbook WriteTextToCell(Workbook workbook, int worksheetPosition, string text, int cellRow = 0, int cellColumn = 0)
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      Cells cells = worksheet.Cells;
      cells[cellRow, cellColumn].PutValue(text);
      return workbook;
    }

    /// <summary>
    /// Set the size of a row.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetPosition">The position of the worksheet to be used.</param>
    /// <param name="rowNumber">The row number to change its size.</param>
    /// <param name="newRowSize">The new size of the row.</param>
    /// <returns>A workbook with a new row size.</returns>
    public static Workbook SetRowSize(Workbook workbook, int worksheetPosition, int rowNumber = 0, double newRowSize = 15)
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      workbook.Worksheets[worksheetPosition].Cells.Rows[rowNumber].Height = newRowSize;
      return workbook;
    }

    /// <summary>
    /// Center the text horizontally and vertically.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetPosition">The position of the worksheet starting with zero.</param>
    /// <param name="styleRowNumber">The number of the row to get the style.</param>
    /// <param name="styleColumnNumber">The number of the column to get the style.</param>
    /// <param name="startingColumn">The column number to start from.</param>
    /// <param name="endingColumn">The column number to end.</param>
    /// <param name="startingRow">The row number to start from.</param>
    /// <param name="endingRow">The row number to end.</param>
    /// <param name="horizontalAlignment">The type of the text alignment horizontally like Left, Right, Center, etc.</param>
    /// <param name="verticalAlignment">The type of the text alignment vertically like Left, Right, Center, etc.</param>
    /// <returns>A workbook with text aligned.</returns>
    public static Workbook SetTextHorizontalAndVertical(Workbook workbook, int worksheetPosition, int styleRowNumber = 0, int styleColumnNumber = 0, int startingColumn = 0, int endingColumn = 0, int startingRow = 0, int endingRow = 0, TextAlignmentType horizontalAlignment = TextAlignmentType.Left, TextAlignmentType verticalAlignment = TextAlignmentType.Center)
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      Cells cells = worksheet.Cells;
      Style style = worksheet.Cells[styleRowNumber, styleColumnNumber].GetStyle();
      style.HorizontalAlignment = horizontalAlignment;
      style.VerticalAlignment = verticalAlignment;

      for (int i = startingRow; i <= endingRow; i++)
      {
        for (int j = startingColumn; j <= endingColumn; j++)
        {
          cells[i, j].SetStyle(style);
        }
      }

      return workbook;
    }

    /// <summary>
    /// Set the background color.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetPosition">The position of the worksheet. The default value is zero.</param>
    /// <param name="styleRowNumber">The row number to get the style from. The default value is zero.</param>
    /// <param name="styleColumnNumber">The column number to get the style from. The default value is zero.</param>
    /// <param name="startingRow">The starting row number to apply the style. The default value is zero.</param>
    /// <param name="endingRow">The ending row number to apply the style. The default value is zero.</param>
    /// <param name="startingColumn">The starting column number to apply the style. The default value is zero.</param>
    /// <param name="endingColumn">The ending column number to apply the style. The default value is zero.</param>
    /// <param name="color">The color of the background.</param>
    /// <returns>A workbook with the requested style applied.</returns>
    public static Workbook SetBackgroundColor(Workbook workbook, Color color, int worksheetPosition = 0, int styleRowNumber = 0, int styleColumnNumber = 0, int startingRow = 0, int endingRow = 0, int startingColumn = 0, int endingColumn = 0)
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      Cells cells = worksheet.Cells;
      Style style = worksheet.Cells[styleRowNumber, styleColumnNumber].GetStyle();
      style.ForegroundColor = color;
      style.Pattern = BackgroundType.Solid;

      for (int i = startingRow; i <= endingRow; i++)
      {
        for (int j = startingColumn; j <= endingColumn; j++)
        {
          cells[i, j].SetStyle(style);
        }
      }

      return workbook;
    }

    /// <summary>
    /// Add a new worksheet at the end of the workbook.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetName">The name of the new worksheet to be inserted.</param>
    /// <returns>A workbook with the new inserted tab.</returns>
    public static Workbook AddTab(Workbook workbook, string worksheetName)
    {
      workbook.Worksheets.Add(worksheetName);
      return workbook;
    }

    /// <summary>
    /// Insert a new worksheet before another one.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="worksheetName">The name of the new worksheet to be inserted.</param>
    /// <param name="tabBeforeNumber">The index to insert a new tab.</param>
    /// <returns>A workbook with the new inserted tab.</returns>
    public static Workbook InsertTab(Workbook workbook, string worksheetName, int tabBeforeNumber = 0)
    {
      var newWorksheet = workbook.Worksheets.Insert(tabBeforeNumber, SheetType.Worksheet);
      newWorksheet.Name = worksheetName;
      return workbook;
    }

    public static string RemoveForbiddencharacters(string projectName)
    {
      string result = string.Empty;
      result = projectName.Remove(' ');
      return result;
    }

    /// <summary>
    /// Remove all Windows forbidden characters for a Windows path.
    /// </summary>
    /// <param name="filename">The initial string to be processed.</param>
    /// <returns>A string without Windows forbidden characters.</returns>
    public static string RemoveWindowsForbiddenCharacters(string filename)
    {
      string result = filename;
      // Remove all characters which are forbidden for a Windows path
      string[] forbiddenWindowsFilenameCharacters = { "\\", "/", "*", "?", "\"", "<", ">", "|" };
      foreach (var item in forbiddenWindowsFilenameCharacters)
      {
        result = result.Replace(item, string.Empty);
      }

      return result;
    }

    /// <summary>
    /// Build a new Excel File and fill it with the structure of a project.
    /// </summary>
    /// <param name="structure">The list of nodes of the structure to be filled.</param>
    /// <param name="sheetName">The name of the worksheet.</param>
    /// <returns>An array of byte with the Excel file.</returns>
    public static byte[] BuildExcelFile(List<StructureExportFormat> structure, string sheetName)
    {
      Dictionary<int, string> dicoPathNode = new Dictionary<int, string>();
      foreach (var item in structure)
      {
        dicoPathNode.Add(item.Header1, item.Header2);
      }

      Workbook workbook = CreateWorkbook(sheetName);
      try
      {
        List<string> headers = new List<string>();
        headers.Add("header1");
        headers.Add("header2");
        headers.Add("header2");

        // Fill the header of the file
        workbook = AddHeader(workbook, headers);

        // Set style for the header
        //todo fix following line
        //workbook = SetStyle(workbook, defaultBlue, Color.White, 12, "Calibri", 0, true, 0, 2, 0, 0, BackgroundType.Solid);

        // import data
        var headersToImport = new string[] { "header1", "header2", "header2" };
        workbook = ImportData(workbook, structure, headersToImport, 0);
        workbook = ReplaceTextWithDictionary(workbook, 0, 1, 2, structure.Count, 2, dicoPathNode);
        workbook = AutoFitColumns(workbook, 0);
        workbook = AddTab(workbook, "Glossary");

      }
      catch (Exception exception)
      {
        throw new Exception("Error while parsing data to export", exception);
      }

      return ConvertSpreadSheetToByteArray(workbook, sheetName);
    }

    private static byte[] ConvertSpreadSheetToByteArray(Workbook workbook, string sheetName)
    {
      throw new NotImplementedException();
    }

    public static Workbook ReplaceTextWithDictionary(Workbook workbook, int worksheetPosition, int startingRow, int startingColumn, int endingRow, int endingColumn, Dictionary<int, string> dicoPathNode)
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      Cells cells = worksheet.Cells;
      for (int i = startingRow; i <= endingRow; i++)
      {
        for (int j = startingColumn; j <= endingColumn; j++)
        {
          string oldCellValue = string.Empty;
          if (cells[i, j].Value == null)
          {
            oldCellValue = string.Empty;
          }
          else
          {
            oldCellValue = cells[i, j].Value.ToString();
          }

          foreach (var item in dicoPathNode.Reverse())
          {
            oldCellValue = oldCellValue.Replace(item.Key.ToString(), item.Value);
            oldCellValue = oldCellValue.Replace("/", " | ");
          }

          cells[i, j].PutValue(oldCellValue.Trim().Trim('|').Trim());
        }
      }

      // delete VIRTUALROOTNODE line
      cells.DeleteRow(dicoPathNode.Count);
      return workbook;
    }

    /// <summary>
    /// Import data to a workbook according to the "Structure Export Format" class.
    /// </summary>
    /// <param name="workbook">The workbook to be used.</param>
    /// <param name="data">The data to be entered.</param>
    /// <param name="worksheetPosition">The position of the worksheet to be modified.</param>
    /// <returns>A workbook with imported data.</returns>
    public static Workbook ImportData(Workbook workbook, List<StructureExportFormat> data, string[] headers, int worksheetPosition = 0)
    {
      Worksheet worksheet = workbook.Worksheets[worksheetPosition];
      Cells cells = worksheet.Cells;

      cells.ImportCustomObjects(
            data.OrderBy(f => f.Header2).ToList(),
            headers,
            false,
            1,
            0,
            data.Count,
            true,
            "dd/mm/yyyy",
            false);

      return workbook;
    }
  }
}
