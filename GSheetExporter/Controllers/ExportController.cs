using GSheetExporter.Models;
using GSheetExporter.Service;
using Microsoft.AspNetCore.Mvc;

namespace GSheetExporter.Controllers
{
    public class ExportController : Controller
    {
        [HttpPost]
        public JsonResult ExportToSheet([FromBody] ExportRequest request)
        {
            var credential = GoogleSheetService.GetCredentialFromJson(request.ServiceAccountJson);

            var data = GoogleSheetService.GenerateMockData();

            var sheetId = request.SpreadSheetId;

            if (string.IsNullOrEmpty(sheetId))
            {
                sheetId = GoogleSheetService.CreateSheet(credential, "User Data");
                if (!string.IsNullOrWhiteSpace(request.ShareWithEmail))
                {
                    GoogleSheetService.ShareWithUser(credential, sheetId, request.ShareWithEmail);
                }

                GoogleSheetService.FillSheet(credential, sheetId, data, false);
            }
            else
            {
                GoogleSheetService.FillSheet(credential, sheetId, data, true);
            }


            var sheetURL = $"https://docs.google.com/spreadsheets/d/{sheetId}/edit";
            return Json(new { url = sheetURL });
        }

        public ActionResult Index()
        {
            return View();
        }
    }
}
