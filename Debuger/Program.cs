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
      //Connecting to sheet
      var sheet1 = new GoogleSheet("1W126R96CXLJJ7x1R14rh70RtGOEEQnJv5J7E32Jx7wI");
      //Creating data and adding it to values
      sheet1.CreateValues(new string[] { "A1","B1","C1" });
      //Updating cells range in tab Balance with created data
      sheet1.UpdateCellsData("Balance", "A1:C1");
      //Reading data from Balance!B1:C1 cells
      sheet1.GetCellsData("Balance", "B1:C1");
      Console.Write(sheet1.values.Range+" -");
      foreach (var item in sheet1.values.Values[0])
      {
        Console.Write(" "+item);
      }
      Console.ReadLine();
    }
  }
}
