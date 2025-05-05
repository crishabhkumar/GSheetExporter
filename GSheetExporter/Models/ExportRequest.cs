namespace GSheetExporter.Models
{
    public class ExportRequest
    {
        public string ServiceAccountJson { get; set; }
        public string ShareWithEmail { get; set; }
        public string SpreadSheetId { get; set; }
    }
}
