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
      var sheet1 = new GoogleSheet("1W126R96CXLJJ7x1R14rh70RtGOEEQnJv5J7E32Jx7wI","Balance");

      //Creating data and adding it to values
      sheet1.CreateValuesLine(new string[] { "A1", "B1", "C1" });

      //Updating cells range in tab Balance with created data
      sheet1.UpdateCellsData("A1:C1");

      //Reading data from Balance!B1:C1 cells
      sheet1.GetCellsData( "B1:C1");
      Console.Write(sheet1.values.Range + " -");
      foreach (var item in sheet1.values.Values[0])
      {
        Console.Write(" " + item);
      }

      //Reading Data from Balance!B1:C1 cells
      sheet1.GetCellsData("A1:C1");
      Console.ReadLine();

      //Adding second line of values as first one comes from Balance!B1:C1
      sheet1.CreateValuesLine(new string[] { "A2", "B2", "C2" });

      //Updating Balance!A1 to C2 range as we have 2 lines of values
      sheet1.UpdateCellsData( "A1");
      Console.ReadLine();
      //cleaning values 
      sheet1.ClearValues();
      //Adding new line of values
      sheet1.CreateValuesLine(new string[] { "1", "2", "3", "4", "5", "6", "7" });
      sheet1.CreateValuesLine(new string[] { "1", "2", "3", "4", "5", "6", "7" });
      //Append data at end of table within range Balance!A1:C1
      sheet1.AppentCellsAtEnd("A1:C1");
      Console.ReadLine();
    }
  }
}
