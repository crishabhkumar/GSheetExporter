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

        public static List<User> GenerateMockData(int count = 1000)
        {
            var users = new List<User>();
            var random = new Random();

            var firstNames = new[] { "John", "Jane", "Alex", "Chris", "Pat", "Taylor", "Sam", "Jamie", "Jordan", "Morgan" };
            var lastNames = new[] { "Doe", "Smith", "Johnson", "Lee", "Walker", "Brown", "Clark", "Miller", "Wilson", "Moore" };
            var domains = new[] { "gmail.com", "yahoo.com", "outlook.com", "example.com" };

            for (int i = 1; i <= count; i++)
            {
                var first = firstNames[random.Next(firstNames.Length)];
                var last = lastNames[random.Next(lastNames.Length)];
                var name = $"{first} {last}";
                var email = $"{first.ToLower()}.{last.ToLower()}{i}@{domains[random.Next(domains.Length)]}";

                users.Add(new User
                {
                    Id = i,
                    Name = name,
                    Email = email
                });
            }

            return users;
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

        public static void FillSheet(GoogleCredential credential, string spreadsheetId, List<User> users, bool startAtLastRow)
        {
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "GSheetExporter"
            });

            int startRow = 1;

            if (startAtLastRow)
            {
                var existing = service.Spreadsheets.Values.Get(spreadsheetId, "Sheet1").Execute();
                startRow = existing.Values?.Count + 1 ?? 1;
                var headers = existing.Values?.FirstOrDefault();

                if (headers == null || !headers.SequenceEqual(new[] { "ID", "Name", "Email" }))
                {
                    throw new Exception("Invalid headers.");
                }
            }

            var values = new List<IList<object>>();

            // Add headers only if we are NOT appending
            if (!startAtLastRow)
            {
                values.Add(new List<object> { "ID", "Name", "Email" });
            }

            values.AddRange(users.Select(u => new List<object> { u.Id, u.Name, u.Email }));

            var valueRange = new ValueRange { Values = values };

            var update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, $"Sheet1!A{startRow}");
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
