using System;

namespace ExcelUtilitiesUsageExample
{
  internal class Program
  {
    static void Main()
    {
      Action<string> Display = Console.WriteLine;
      Display("Example of usage of Excel Utilities library");


      Display("Press any key to exit:");
      Console.ReadLine();
    }
  }
}
