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
            var sheetId = GoogleSheetService.CreateSheet(credential, "User Data");
            GoogleSheetService.FillSheet(credential, sheetId, data);
            GoogleSheetService.ShareWithUser(credential, sheetId, request.ShareWithEmail);

            var sheetURL = $"https://docs.google.com/spreadsheets/d/{sheetId}/edit";
            return Json(new { url = sheetURL });
        }

        public ActionResult Index()
        {
            return View();
        }
    }
}
