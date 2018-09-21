using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SheetTools

{
  class Program
  {
    static void Main(string[] args)
    {
      var sheet1 = new GoogleSheet("1W126R96CXLJJ7x1R14rh70RtGOEEQnJv5J7E32Jx7wI");
      var result = sheet1.GetCellData("Balance", "A", "1");
      Console.WriteLine($"{sheet1.GetCellData("Balance", "B", "1")}+{sheet1.GetCellData("Balance","C","1")}={result}");
      sheet1.RawUpdate(sheet1.GenerateData(new string[1]{"A1"}), "A1");
      Console.WriteLine($"{sheet1.GetCellData("Balance", "B", "1")}+{sheet1.GetCellData("Balance", "C", "1")}={result}");
      Console.ReadLine();
    }
  }
}
