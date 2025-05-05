using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using GSheetExporter.Models;
using System.Text;

namespace GSheetExporter.Service
{
    public static class GoogleSheetService
    {
        public static GoogleCredential GetCredentialFromJson(string serviceAccountJson)
        {
            var byteArray = Encoding.UTF8.GetBytes(serviceAccountJson);
            using (var stream = new MemoryStream(byteArray))
            {
                return GoogleCredential.FromStream(stream).CreateScoped(new[]
                {
                    SheetsService.Scope.Spreadsheets,
                    DriveService.Scope.Drive
                });
            }
        }

        public static List<User> GenerateMockData()
        {
            return new List<User>() {
                new User() {
                    Id = 1,
                    Name = "John Doe",
                    Email = "john@gmail.com" },
                new User() {
                    Id = 2,
                    Name = "Jane Smith",
                    Email = "jane@gmail.com"
                    },
            };
        }

        public static string CreateSheet(GoogleCredential credential, string title)
        {
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "GSheetExporter"
            });

            var sheet = new Spreadsheet { Properties = new SpreadsheetProperties { Title = title } };

            var create = service.Spreadsheets.Create(sheet);
            var response = create.Execute();
            return response.SpreadsheetId;
        }

        public static void FillSheet(GoogleCredential credential, string spreadSheetId, List<User> users)
        {
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "GSheetExporter"
            });

            var values = new List<IList<object>>
            {
                new List<object> { "ID", "Name", "Email" }
            };


            values.AddRange(users.Select(u => new List<object> { u.Id, u.Name, u.Email }));

            var valueRange = new ValueRange { Values = values };

            var update = service.Spreadsheets.Values.Update(valueRange, spreadSheetId, "Sheet1!A1");
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            update.Execute();
        }

        public static void ShareWithUser(GoogleCredential credential, string spreadSheetId, string email)
        {
            var driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "GSheetExporter"
            });
            var permission = new Google.Apis.Drive.v3.Data.Permission
            {
                Type = "user",
                Role = "writer",
                EmailAddress = email
            };
            var request = driveService.Permissions.Create(permission, spreadSheetId);
            request.Fields = "id";
            request.SendNotificationEmail = true;
            request.Execute();
        }

    }
}
