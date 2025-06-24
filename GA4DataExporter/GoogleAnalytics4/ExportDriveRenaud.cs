using System;
using System.Collections.Generic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using Newtonsoft.Json;
using Aspose.Cells;

namespace ExcelToGoogle
{
    public static class ExportDriveRenaud
    {
        public static SheetsService ConnectToGoogle()
        {
            string[] Scopes = { SheetsService.Scope.Spreadsheets };
            string ApplicationName = "Excel to Google Sheet";

            UserCredential credential;

            using (var stream = new FileStream("oog-stagiaire-sheet-report.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            var result =  new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return result;
        }

        public static void ExportDataFromExcelToGoogleSheet(SheetsService sheetService, string excelFileName)
        {
            if (!File.Exists(excelFileName))
            {
                Console.WriteLine($"Fichier Excel introuvable: {excelFileName}");
                return;
            }

            Workbook wb = new Workbook(excelFileName);
            string defaultWorksheetName = wb.Worksheets[0].Name;
            Spreadsheet spreadsheet = CreateSpreadsheet(sheetService, Path.GetFileNameWithoutExtension(excelFileName), defaultWorksheetName);

            Console.WriteLine($"Spreadsheet créée: {spreadsheet.SpreadsheetUrl}");
            Console.WriteLine($"ID: {spreadsheet.SpreadsheetId}");

            foreach (var sheet in wb.Worksheets)
            {
                string range = $"{sheet.Name}!A:Z";

                if (sheet.Index > 0)
                {
                    AddSheet(sheetService, spreadsheet.SpreadsheetId, sheet.Name);
                }

                IList<IList<object>> data = new List<IList<object>>();
                for (int i = 0; i <= sheet.Cells.MaxDataRow; i++)
                {
                    List<object> rowData = new List<object>();
                    for (int j = 0; j <= sheet.Cells.MaxDataColumn; j++)
                    {
                        rowData.Add(sheet.Cells[i, j].Value?.ToString() ?? "");
                    }
                    data.Add(rowData);
                }

                UpdateSpreadsheet(sheetService, spreadsheet.SpreadsheetId, range, data);
            }
        }

        private static Spreadsheet CreateSpreadsheet(SheetsService sheetsService, string spreadsheetName, string defaultSheetName)
        {
            return sheetsService.Spreadsheets.Create(new Spreadsheet()
            {
                Properties = new SpreadsheetProperties() { Title = spreadsheetName },
                Sheets = new List<Sheet>() { new Sheet() { Properties = new SheetProperties() { Title = defaultSheetName } } }
            }).Execute();
        }

        private static void AddSheet(SheetsService sheetsService, string spreadSheetId, string sheetName)
        {
            sheetsService.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>() { new Request() { AddSheet = new AddSheetRequest() { Properties = new SheetProperties() { Title = sheetName } } } }
            }, spreadSheetId).Execute();
        }

        private static void UpdateSpreadsheet(SheetsService sheetService, string spreadsheetId, string range, IList<IList<object>> data)
        {
            var valueRange = new ValueRange() { Range = range, Values = data };
            var updateRequest = sheetService.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateRequest.Execute();
        }
    }
}
