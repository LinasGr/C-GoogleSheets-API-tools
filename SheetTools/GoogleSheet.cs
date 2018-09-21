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
    string[] Scopes = { SheetsService.Scope.Spreadsheets };
    string applicationName = "CSharp";
    String sheetID;
    SheetsService service;
    public String range;
    public ValueRange values;

    public GoogleSheet(string spreadSheetID)
    {
      this.sheetID = spreadSheetID;
      this.service = AuthorizeGoogleApp();
      this.range = "A1";
      this.values = new ValueRange();
      this.values.Values = new List<IList<object>>();
      this.values.Values.Add(new List<object>());
    }

    public GoogleSheet()
    {
      this.sheetID = "";
      this.service = AuthorizeGoogleApp();
      this.range = "A1";
      this.values = new ValueRange();
      this.values.Values = new List<IList<object>>();
      this.values.Values.Add(new List<object>());
    }

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

    public void GetCellsData(string tab, string range)
    {
      this.range = tab + "!" + range;
      this.values.Range = this.range;
      SpreadsheetsResource.ValuesResource.GetRequest request = this.service.Spreadsheets.Values.Get(this.sheetID, this.range);
      request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMATTEDVALUE;
      this.values = request.Execute();
    }

    public void UpdateCellsData(string tab, string range)
    {
      this.range = tab + "!" + range;
      this.values.Range = this.range;
      SpreadsheetsResource.ValuesResource.UpdateRequest request =
        service.Spreadsheets.Values.Update(this.values, this.sheetID, this.range);
      request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
      request.Execute();
    }

    public void ClearValues()
    {
      List<IList<Object>> objNewRecords = new List<IList<Object>>();
      this.values.Values = objNewRecords;
      this.values.Values.Add(new List<Object>());
    }
    public void CreateValues(string[] arr)
    {
      List<IList<Object>> objNewRecords = new List<IList<Object>>();

      IList<Object> obj = new List<Object>();
      foreach (var item in arr)
      {
        obj.Add(item);
        this.values.Values.Add(new List<object>(){item});
      }
      objNewRecords.Add(obj);

      this.values.Values = objNewRecords;
    }

    public void RawAppent(IList<IList<Object>> values, string newRange)
    {
      SpreadsheetsResource.ValuesResource.AppendRequest request =
         this.service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, this.sheetID, newRange);
      request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
      request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
      var response = request.Execute();
    }
    public void RawUpdate(IList<IList<object>> values, string newRange)
    {
      SpreadsheetsResource.ValuesResource.UpdateRequest request =
        this.service.Spreadsheets.Values.Update(new ValueRange() { Values = values }, this.sheetID, newRange);
      request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
      var response = request.Execute();
    }
  }
}
