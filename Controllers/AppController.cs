using System.Diagnostics;
using Markerting.Models;
using Markerting.Services;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace Markerting.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly GenericExcelAnalysisService _genericExcel;

        public HomeController(ILogger<HomeController> logger, GenericExcelAnalysisService genericExcel)
        {
            _logger = logger;
            _genericExcel = genericExcel;
        }

        public IActionResult Index()
        {
            var session = HttpContext.Session.GetString("GenericExcelDataset");
            var payload = GenericExcelAnalysisService.DeserializePayload(session);
            if (payload == null)
                return View(new GenericDashboardViewModel { HasData = false });

            var reqJson = HttpContext.Session.GetString("GenericAnalysisRequest");
            var request = GenericExcelAnalysisService.DeserializeAnalysisRequest(reqJson);
            if (request == null)
                return RedirectToAction(nameof(AnalyzeSetup));

            var dashboard = _genericExcel.BuildDashboard(payload, request);
            return View(dashboard);
        }

        public IActionResult AnalyzeSetup()
        {
            var session = HttpContext.Session.GetString("GenericExcelDataset");
            var payload = GenericExcelAnalysisService.DeserializePayload(session);
            if (payload == null)
            {
                TempData["ErrorMessage"] = "Upload a workbook first.";
                return RedirectToAction(nameof(Index));
            }

            var reqJson = HttpContext.Session.GetString("GenericAnalysisRequest");
            var current = GenericExcelAnalysisService.DeserializeAnalysisRequest(reqJson);
            var vm = _genericExcel.BuildAnalyzeSetupViewModel(payload, current);
            return View(vm);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult RunAnalysis([FromForm] AnalyzeSetupFormModel model)
        {
            var session = HttpContext.Session.GetString("GenericExcelDataset");
            var payload = GenericExcelAnalysisService.DeserializePayload(session);
            if (payload == null)
            {
                TempData["ErrorMessage"] = "Your upload expired. Please upload again.";
                return RedirectToAction(nameof(Index));
            }

            var err = _genericExcel.ValidateAnalysisForm(payload, model, out var request);
            if (err != null)
            {
                TempData["ErrorMessage"] = err;
                var vm = _genericExcel.BuildAnalyzeSetupViewModel(payload, MapFormToRequest(model));
                return View(nameof(AnalyzeSetup), vm);
            }

            HttpContext.Session.SetString("GenericAnalysisRequest", GenericExcelAnalysisService.SerializeAnalysisRequest(request));
            TempData["SuccessMessage"] = "Analysis settings saved. Here is your dashboard.";
            return RedirectToAction(nameof(Index));
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["ErrorMessage"] = "Please select a file to upload.";
                return RedirectToAction(nameof(Index));
            }

            if (!file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                !file.FileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            {
                TempData["ErrorMessage"] = "Please upload an Excel file (.xlsx recommended).";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                await using var readStream = new MemoryStream();
                await file.CopyToAsync(readStream);
                readStream.Position = 0;

                var payload = _genericExcel.ParseUpload(readStream, file.FileName);
                if (payload == null || payload.Rows.Count == 0)
                {
                    TempData["ErrorMessage"] = "Could not read the workbook. Use the first sheet with a header row and at least one data row (.xlsx works best).";
                    return RedirectToAction(nameof(Index));
                }

                var json = GenericExcelAnalysisService.SerializePayload(payload);
                HttpContext.Session.SetString("GenericExcelDataset", json);
                HttpContext.Session.Remove("GenericAnalysisRequest");

                var note = payload.WasTruncated
                    ? $" Analysis uses the first {GenericExcelAnalysisService.MaxRowsStored:N0} non-empty rows for performance."
                    : string.Empty;

                TempData["SuccessMessage"] =
                    $"Loaded \"{payload.FileName}\" ({payload.Rows.Count:N0} rows, {payload.Headers.Count} columns) from sheet \"{payload.SheetName}\".{note} Choose what to analyze next.";
                return RedirectToAction(nameof(AnalyzeSetup));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing Excel file");
                TempData["ErrorMessage"] = $"Could not read that file. Try saving as .xlsx. Details: {ex.Message}";
                return RedirectToAction(nameof(Index));
            }
        }

        public IActionResult ExportGeneric()
        {
            var session = HttpContext.Session.GetString("GenericExcelDataset");
            var payload = GenericExcelAnalysisService.DeserializePayload(session);
            if (payload == null || payload.Headers.Count == 0)
            {
                TempData["ErrorMessage"] = "Upload a workbook before exporting.";
                return RedirectToAction(nameof(Index));
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage();
            var dataSheet = package.Workbook.Worksheets.Add("Data");
            for (int c = 0; c < payload.Headers.Count; c++)
            {
                dataSheet.Cells[1, c + 1].Value = payload.Headers[c];
                dataSheet.Cells[1, c + 1].Style.Font.Bold = true;
            }

            var rowIndex = 2;
            foreach (var row in payload.Rows)
            {
                for (int c = 0; c < payload.Headers.Count; c++)
                    dataSheet.Cells[rowIndex, c + 1].Value = c < row.Count ? row[c] : string.Empty;
                rowIndex++;
            }

            dataSheet.Cells.AutoFitColumns();

            var profile = package.Workbook.Worksheets.Add("Column profile");
            var reqJson = HttpContext.Session.GetString("GenericAnalysisRequest");
            var analysisReq = GenericExcelAnalysisService.DeserializeAnalysisRequest(reqJson);
            var dash = _genericExcel.BuildDashboard(payload, analysisReq);
            profile.Cells[1, 1].Value = "Column";
            profile.Cells[1, 2].Value = "Type";
            profile.Cells[1, 3].Value = "Non-empty";
            profile.Cells[1, 4].Value = "Notes";
            profile.Cells[1, 1, 1, 4].Style.Font.Bold = true;

            var profileRow = 2;
            foreach (var col in dash.Columns)
            {
                profile.Cells[profileRow, 1].Value = col.Name;
                profile.Cells[profileRow, 2].Value = col.Kind.ToString();
                profile.Cells[profileRow, 3].Value = col.NonEmptyCount;
                var notes = col.Kind == GenericColumnKind.Number
                    ? $"Sum {col.Sum:N2}, Avg {col.Average:N2}"
                    : col.Kind == GenericColumnKind.Date
                        ? $"{col.DateMin} -> {col.DateMax}"
                        : $"~{col.DistinctApprox} distinct values";
                profile.Cells[profileRow, 4].Value = notes;
                profileRow++;
            }

            profile.Cells.AutoFitColumns();

            var stream = new MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;
            var safe = Path.GetFileNameWithoutExtension(payload.FileName);
            if (string.IsNullOrWhiteSpace(safe))
                safe = "export";

            var fileName = $"{safe}_export_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        private static UserAnalysisRequest MapFormToRequest(AnalyzeSetupFormModel form) =>
            new()
            {
                Goal = form.Goal,
                CategoryColumn = form.CategoryColumn,
                MeasureColumn = form.MeasureColumn,
                DateColumn = form.DateColumn,
                CompareValueA = form.CompareValueA,
                CompareValueB = form.CompareValueB
            };
    }
}
