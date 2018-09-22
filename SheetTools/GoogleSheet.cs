using System;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Collections.Generic;

namespace SheetTools
{
  public class GoogleSheet
  {
    //Scopes - permitions working with SpreadSheets
    string[] Scopes = { SheetsService.Scope.Spreadsheets };
    //Name of aplications for statistics on google API
    string applicationName = "CSharp";
    //spreadSheetID - key to specific sheet
    String sheetID;

    SheetsService service;

    //range is address of tab and cell or range of cells
    //required when fetching from or pushing data to sheets
    public string range;

    //Tab is sheet name object is working with
    public string tab;

    //values keeps data got from sheet or ready to be pushed to sheets
    public ValueRange values;

    //spreadSheetID - key to specific sheet
    //tab is name of sheet object works on
    public GoogleSheet(string spreadSheetID,string tab)
    {
      this.sheetID = spreadSheetID;
      this.service = AuthorizeGoogleApp();
      this.tab = tab;
      this.range = "A1";//On creation points to left top corner of Sheet
      this.values = new ValueRange();
      this.values.Values = new List<IList<object>>();
      this.values.Values.Add(new List<object>());
    }

    //Method for constructor to conecct sheet
    private SheetsService AuthorizeGoogleApp()
    {
      UserCredential credential;
      using (var stream =
          new FileStream("client_id.json", FileMode.Open, FileAccess.Read))
      {
        string credPath = System.Environment.GetFolderPath(
            System.Environment.SpecialFolder.Personal);
        credPath = "token.json";
        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
            GoogleClientSecrets.Load(stream).Secrets,
            this.Scopes,
            "user",
            CancellationToken.None,
            new FileDataStore(credPath, true)).Result;
        Console.WriteLine("Credential file saved to: " + credPath);
      }
      // Create Google Sheets API service.
      var service = new SheetsService(new BaseClientService.Initializer()
      {
        HttpClientInitializer = credential,
        ApplicationName = applicationName,
      });
      return service;
    }

    //Reads data to values from  tab!range
    public void GetCellsData( string range)
    {
      this.range = this.tab + "!" + range;
      this.values.Range = this.range;
      SpreadsheetsResource.ValuesResource.GetRequest request = this.service.Spreadsheets.Values.Get(this.sheetID, this.range);
      request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMATTEDVALUE;
      this.values = request.Execute();
    }

    //Updates sheet at tab!range with data from values
    public void UpdateCellsData(string range)
    {
      this.range = this.tab + "!" + range;
      this.values.Range = this.range;
      SpreadsheetsResource.ValuesResource.UpdateRequest request =
        service.Spreadsheets.Values.Update(this.values, this.sheetID, this.range);
      request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
      request.Execute();
    }

    //resets values
    public void ClearValues()
    {
      List<IList<Object>> objNewRecords = new List<IList<Object>>();
      this.values.Values = objNewRecords;
      this.values.Values.Add(new List<Object>());
    }

    //Adds line of data to values
    public void CreateValuesLine(string[] arr)
    {
      IList<Object> obj = new List<Object>();
      foreach (var item in arr)
      {
        obj.Add(item);
      }
      if (this.values.Values.Count == 1)
      {
        if (this.values.Values[0].Count == 0) this.values.Values[0] = obj;
        else this.values.Values.Add(obj);
      }
      else this.values.Values.Add(obj);
    }

    //Appends values at the and of table 
    public void AppentCellsAtEnd(string range)
    {
      this.range = this.tab+"!"+range;
      this.values.Range = this.range;
      SpreadsheetsResource.ValuesResource.AppendRequest request =
         this.service.Spreadsheets.Values.Append(this.values, this.sheetID, this.range);
      request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
      request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
      var response = request.Execute();
    }
  }
}
