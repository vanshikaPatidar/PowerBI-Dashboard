using Microsoft.AspNetCore.Mvc;
using PowerBI.Models;
using PowerBI.Services;
using PowerBI.Data;
using System.Diagnostics;

namespace PowerBI.Controllers
{
    public class ReportController : Controller
    {
        private readonly PowerBIService _service;
        private readonly AppDbContext _db;

        public ReportController(PowerBIService service, AppDbContext db)
        {
            _service = service;
            _db = db;
        }

        public async Task<IActionResult> Index(int workspaceId)
        {
            Console.WriteLine($"[REPORT] Listing reports for Local Workspace ID: {workspaceId}");
            
            var ws = _db.Workspaces.Find(workspaceId);
            if (ws == null || string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) 
                return RedirectToAction("Index", "Workspace");

            try
            {
                // Sync reports within this workspace
                await _service.SyncReports(workspaceId, Guid.Parse(ws.PowerBIWorkspaceId), _db);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[REPORT] Sync Error: {ex.Message}");
                ViewBag.SyncError = "Failed to sync reports: " + ex.Message;
            }

            var reports = _db.Reports
                .Where(r => r.WorkspaceId == workspaceId)
                .ToList();

            ViewBag.WorkspaceId = workspaceId;
            ViewBag.WorkspaceName = ws.Name;
            Console.WriteLine($"[REPORT] Returning {reports.Count} synced reports.");

            return View(reports);
        }

        [HttpPost]
        public async Task<IActionResult> Upload(IFormFile file, int workspaceId)
        {
            Console.WriteLine($"[REPORT] Uploading file: {file.FileName} to Workspace: {workspaceId}");
            
            var ws = _db.Workspaces.Find(workspaceId);
            if (ws == null || string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) 
            {
                Console.WriteLine("[REPORT] ERROR: Workspace not found.");
                return BadRequest("Invalid Workspace");
            }

            var uploadsDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
            if (!Directory.Exists(uploadsDir)) Directory.CreateDirectory(uploadsDir);
            var path = Path.Combine(uploadsDir, file.FileName);

            try 
            {
                using var stream = file.OpenReadStream();
                var import = await _service.UploadReport(
                    Guid.Parse(ws.PowerBIWorkspaceId),
                    file.FileName,
                    stream);

                // Save locally after successful (or attempted) upload for record keeping
                using (var fileStream = new FileStream(path, FileMode.Create))
                {
                    await file.OpenReadStream().CopyToAsync(fileStream);
                }

                if (import.Reports == null || !import.Reports.Any())
                {
                    Console.WriteLine("[REPORT] ERROR: Power BI Import returned no reports.");
                    return BadRequest("Import failed or no reports found.");
                }

                var reportId = import.Reports.First().Id;
                Console.WriteLine($"[REPORT] Import successful. Power BI Report ID: {reportId}");

                _db.Reports.Add(new Report
                {
                    Name = file.FileName,
                    PowerBIReportId = reportId.ToString(),
                    WorkspaceId = workspaceId,
                    FilePath = path,
                    ReportType = file.FileName.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase) ? "RDL" : "PowerBI"
                });

                await _db.SaveChangesAsync();
                Console.WriteLine("[REPORT] Record saved to local database.");

                return RedirectToAction("Index", new { workspaceId });
            }
            catch (Exception ex)
            {
                var message = ex.Message;
                if (message.Contains("Unauthorized"))
                {
                    message = "Unauthorized. Please ensure the Service Principal (App Registration) is added as an ADMIN or MEMBER to this specific workspace in the Power BI portal.";
                }
                else if (message.Contains("Forbidden"))
                {
                    message = "Forbidden. This usually means the Service Principal is not an admin of the workspace, or 'Export to PDF' is disabled for SPs in Tenant settings.";
                }
                else if (message.Contains("RequestedFileIsEncryptedOrCorrupted"))
                {
                    bool isRdl = file.FileName.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase);
                    message = "The file appears to be corrupted, encrypted, or not a valid Power BI report.";
                    if (isRdl)
                    {
                        message += " IMPORTANT: Paginated Reports (.rdl) require the target workspace to be on a Premium or Fabric capacity.";
                    }
                    message += " Also ensure the file is not protected by sensitivity labels.";
                }
                TempData["Error"] = "Upload Error: " + message;
                return RedirectToAction("Index", new { workspaceId });
            }
        }

        public async Task<IActionResult> Export(int reportId)
        {
            Console.WriteLine($"[REPORT] Downloading PDF for Report ID: {reportId}");
            
            var report = _db.Reports.Find(reportId);
            if (report == null || string.IsNullOrEmpty(report.PowerBIReportId)) return NotFound();

            var ws = _db.Workspaces.Find(report.WorkspaceId);
            if (ws == null || string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) return BadRequest("Invalid Workspace");

            var pdfStream = await _service.ExportReportAsStream(
                Guid.Parse(ws.PowerBIWorkspaceId),
                Guid.Parse(report.PowerBIReportId));

            var fileName = $"{report.Name}_{DateTime.Now:yyyyMMddHHmmss}.pdf";
            var uploadsDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
            var filePath = Path.Combine(uploadsDir, fileName);

            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                await pdfStream.CopyToAsync(fileStream);
            }

            Console.WriteLine($"[REPORT] PDF saved to local server: {filePath}");
            
            var bytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(bytes, "application/pdf", fileName);
        }

        public async Task<IActionResult> Preview(int reportId)
        {
            Console.WriteLine($"[REPORT] Generating Preview for Report ID: {reportId}");
            
            var report = _db.Reports.Find(reportId);
            if (report == null || string.IsNullOrEmpty(report.PowerBIReportId)) return NotFound();

            var ws = _db.Workspaces.Find(report.WorkspaceId);
            if (ws == null || string.IsNullOrEmpty(ws.PowerBIWorkspaceId)) return BadRequest("Invalid Workspace");

            var embedConfig = await _service.GetEmbedConfig(
                Guid.Parse(ws.PowerBIWorkspaceId),
                Guid.Parse(report.PowerBIReportId));
            
            embedConfig.ReportType = report.ReportType;
            embedConfig.LocalReportId = report.Id;
            embedConfig.WorkspaceId = ws.PowerBIWorkspaceId;

            Console.WriteLine($"[REPORT] Embed Config generated for: {report.Name} ({report.ReportType})");
            return View(embedConfig);
        }

        [HttpPost]
        public async Task<IActionResult> ResetFilters(int reportId)
        {
            var report = _db.Reports.Find(reportId);
            if (report == null) return NotFound();

            // 1. Purge all existing metadata
            var existing = _db.ReportFilters.Where(f => f.ReportId == reportId).ToList();
            Console.WriteLine($"[CONTROLLER] EXPLICIT RESET: DELETING {existing.Count} ENTRIES FOR REPORT {reportId}");
            
            _db.ReportFilters.RemoveRange(existing);
            await _db.SaveChangesAsync();

            // 2. Force fresh discovery
            try
            {
                Guid? dsId = !string.IsNullOrEmpty(report.PowerBIDatasetId) ? Guid.Parse(report.PowerBIDatasetId) : (Guid?)null;
                await _service.DiscoverReportFilters(reportId, dsId, _db);
                return Ok(new { message = "Schema discovered and saved successfully." });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[CONTROLLER] Discovery failed: {ex.Message}");
                return BadRequest(new { message = $"Discovery failed: {ex.Message}" });
            }
        }

        [HttpGet]
        public async Task<IActionResult> GetFilters(int reportId)
        {
            var filters = _db.ReportFilters
                .Where(f => f.ReportId == reportId && f.IsActive)
                .ToList();
            
            return Json(filters);
        }

        [HttpPost]
        public async Task<IActionResult> SaveManualFilter(int reportId, string displayName, string tableName, string columnName)
        {
            if (string.IsNullOrEmpty(tableName) || string.IsNullOrEmpty(columnName))
                return BadRequest("Table and Column names are required.");

            var filter = new ReportFilter
            {
                ReportId = reportId,
                DisplayName = displayName ?? columnName,
                TableName = tableName,
                ColumnName = columnName,
                IsActive = true
            };

            _db.ReportFilters.Add(filter);
            await _db.SaveChangesAsync();

            return Ok(new { message = "Filter added successfully." });
        }

        [HttpGet]
        public async Task<IActionResult> GetFilterValues(string? datasetId, string table, string column, string? reportId, string? workspaceId)
        {
            if (string.IsNullOrEmpty(table) || string.IsNullOrEmpty(column))
                return BadRequest("Missing parameters");

            Guid? dsGuid = !string.IsNullOrEmpty(datasetId) ? Guid.Parse(datasetId) : (Guid?)null;
            Guid? repGuid = !string.IsNullOrEmpty(reportId) ? Guid.Parse(reportId) : (Guid?)null;
            Guid? wsGuid = !string.IsNullOrEmpty(workspaceId) ? Guid.Parse(workspaceId) : (Guid?)null;

            var values = await _service.GetColumnValues(dsGuid, table, column, repGuid, wsGuid);
            return Json(values);
        }

    }
}
