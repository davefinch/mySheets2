using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace SheetsQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets};
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = "1Nmd2kup-kx5KANlGz0r7esH_lEDahhrJBVG-eRwynB8";
            String range = "Sheet1!A2:E";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);


            // Define request parameters.
            //https://docs.google.com/spreadsheets/d/171wIpyVxTFVYpHs1j7GpBUlnIXGxqMpPMdMOMx5w1IE/edit?usp=sharing;
            //spreadsheetId = "1Nmd2kup-kx5KANlGz0r7esH_lEDahhrJBVG-eRwynB8";

            var sheet = "Sheet1";

            range = $"{sheet}!A:F";
            var valueRange = new ValueRange();

            var oblist = new List<object>() { "Hello!", "This", "was", "insertd", "too", "weds" };
            valueRange.Values = new List<IList<object>> { oblist };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = appendRequest.Execute();


            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}", row[0], row[4]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Console.Read();
        }
    }
}



//using Google.Apis.Auth.OAuth2;
//using Google.Apis.Sheets.v4;
//using Google.Apis.Sheets.v4.Data;
//using Google.Apis.Services;
//using Google.Apis.Util.Store;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Threading;

//namespace mySheets2
//{
//    class Program
//    {
//        // If modifying these scopes, delete your previously saved credentials
//        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
//        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
//        static string ApplicationName = "Google Sheets API .NET Quickstart";

//        static void Main(string[] args)
//        {
//            UserCredential credential;

//            using (var stream =
//                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
//            {
//                // The file token.json stores the user's access and refresh tokens, and is created
//                // automatically when the authorization flow completes for the first time.
//                string credPath = "token.json";
//                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
//                    GoogleClientSecrets.Load(stream).Secrets,
//                    Scopes,
//                    "user",
//                    CancellationToken.None,
//                    new FileDataStore(credPath, true)).Result;
//                Console.WriteLine("Credential file saved to: " + credPath);
//            }

//            // Create Google Sheets API service.
//            var service = new SheetsService(new BaseClientService.Initializer()
//            {
//                HttpClientInitializer = credential,
//                ApplicationName = ApplicationName,
//            });

//            // Define request parameters.
//            //https://docs.google.com/spreadsheets/d/171wIpyVxTFVYpHs1j7GpBUlnIXGxqMpPMdMOMx5w1IE/edit?usp=sharing;
//            //String spreadsheetId = "1Nmd2kup-kx5KANlGz0r7esH_lEDahhrJBVG-eRwynB8";

//            //var sheet = "Sheet1";

//            //var range = $"{sheet}!A:F";
//            //var valueRange = new ValueRange();

//            //var oblist = new List<object>() { "Hello!", "This", "was", "insertd", "via", "C#" };
//            //valueRange.Values = new List<IList<object>> { oblist };

//            //var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
//            //appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
//            //var appendReponse = appendRequest.Execute();




//            //    SpreadsheetsResource.BatchUpdateRequest req = service.Spreadsheets.BatchUpdate((vr, spreadsheetId, range);


//            // Create Google Sheets API service.
//            var service = new SheetsService(new BaseClientService.Initializer()
//            {
//                HttpClientInitializer = credential,
//                ApplicationName = ApplicationName,
//            });

//            // Define request parameters.
//            String spreadsheetId = "1Nmd2kup-kx5KANlGz0r7esH_lEDahhrJBVG-eRwynB8";
//            String range = "Class Data!A2:E";
//            SpreadsheetsResource.ValuesResource.GetRequest request =
//                    service.Spreadsheets.Values.Get(spreadsheetId, range);

//            // Prints the names and majors of students in a sample spreadsheet:
//            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
//            ValueRange response = request.Execute();
//            IList<IList<Object>> values = response.Values;
//            if (values != null && values.Count > 0)
//            {
//                Console.WriteLine("Name, Major");
//                foreach (var row in values)
//                {
//                    // Print columns A and E, which correspond to indices 0 and 4.
//                    Console.WriteLine("{0}, {1}", row[0], row[4]);
//                }
//            }
//            else
//            {
//                Console.WriteLine("No data found.");
//            }
//            Console.Read();
//        }
//    }
//}