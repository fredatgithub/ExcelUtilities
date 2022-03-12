using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Aspose;

namespace ExcelUtilities.Tests
{
  [TestClass]
  public class UnitTestExcelMethods
  {
    [TestMethod]
    public void TestMethod_Create_Workbook_Is_Not_null()
    {
      var source = ExcelMethods.CreateWorkbook();
      Assert.IsNotNull(source);
    }

    [TestMethod]
    public void TestMethod_2()
    {
      var source = ExcelMethods.CreateWorkbook();
      var source2 = source.Worksheets;
      Assert.IsNotNull(source2);
    }
  }
}
