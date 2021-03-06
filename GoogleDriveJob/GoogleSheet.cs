﻿using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using GoogleDriveJob.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GoogleDriveJob
{
    public class GoogleSheet
    {
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private const string GoogleCredentialsFileName = "sheet-credentials.json";
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        private static SheetsService GetSheetsService()
        {
            UserCredential credential;

            using (var stream =
                new FileStream(GoogleCredentialsFileName, FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    $"{Environment.UserName}",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        public static async Task WriteAsync(List<IList<Object>> values, string spreadSheetId, string fileName)
        {
            var sheetService = GetSheetsService();
            var valueRange = new ValueRange()
            {
                MajorDimension = "ROWS",
                Values = values
            };

            // How the input data should be interpreted.
            SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum valueInputOption = (SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED);  // TODO: Update placeholder value.

            // How the input data should be inserted.
            SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum insertDataOption = (SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS);  // TODO: Update placeholder value.

            SpreadsheetsResource.ValuesResource.AppendRequest request = sheetService.Spreadsheets.Values.Append(valueRange, spreadSheetId, "A:A");
            request.ValueInputOption = valueInputOption;
            request.InsertDataOption = insertDataOption;

            AppendValuesResponse response = await request.ExecuteAsync();

            if (response?.Updates?.UpdatedRows != null)
            {
                Console.WriteLine($"{response.Updates.UpdatedRows} Rows Updated for Sheet: {fileName}");
            }
            else
            {
                Console.WriteLine($"Sheet : {fileName} is up to date");
            }
        }

        protected static string GetRange(SheetsService service, string SheetId)
        {
            // Define request parameters.  
            String spreadsheetId = SheetId;
            String range = "A:A";

            SpreadsheetsResource.ValuesResource.GetRequest getRequest =
                       service.Spreadsheets.Values.Get(spreadsheetId, range);
            System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
            ValueRange getResponse = getRequest.Execute();
            IList<IList<Object>> getValues = getResponse.Values;
            if (getValues == null)
            {
                // spreadsheet is empty return Row A Column A  
                return "A:A";
            }

            int currentCount = getValues.Count() + 1;
            String newRange = "A" + currentCount + ":A";
            return newRange;
        }
    }
}
