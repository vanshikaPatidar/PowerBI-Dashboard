using Microsoft.PowerBI.Api;
using Microsoft.PowerBI.Api.Models;
using Microsoft.Rest;
using PowerBI.Models;
using PowerBI.Data;
using System.Net.Http.Headers;
using System.Linq;
using System.Collections.Generic;

namespace PowerBI.Services
{
    public class PowerBIService
    {
        private readonly PowerBIAuthService _auth;

        public PowerBIService(PowerBIAuthService auth)
        {
            _auth = auth;
        }

        private async Task<PowerBIClient> GetClient()
        {
            Console.WriteLine("[SERVICE] Fetching token from Azure AD...");
            var token = await _auth.GetAccessToken();
            
            // Log first 10 chars of token for debugging (Safe)
            Console.WriteLine($"[SERVICE] Token acquired (Starts with: {token.Substring(0, 10)}...)");

            var credentials = new TokenCredentials(token, "Bearer");
            return new PowerBIClient(new Uri("https://api.powerbi.com/"), credentials);
        }

        // 🏢 Get or Create Workspace
        public async Task<Microsoft.PowerBI.Api.Models.Group> GetOrCreateWorkspace(string name)
        {
            try 
            {
                Console.WriteLine($"[SERVICE] Checking if workspace '{name}' exists...");
                var client = await GetClient();
                var groups = await client.Groups.GetGroupsAsync();
                
                var existingGroup = groups?.Value?.FirstOrDefault(g => g.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                
                if (existingGroup != null)
                {
                    Console.WriteLine($"[SERVICE] Found existing workspace: {existingGroup.Id}");
                    return existingGroup;
                }

                Console.WriteLine("[SERVICE] Not found. Creating new workspace...");
                var newGroup = await client.Groups.CreateGroupAsync(new GroupCreationRequest { Name = name });
                Console.WriteLine($"[SERVICE] New workspace created: {newGroup.Id}");

                // 🚀 AUTO-ASSIGN TO CAPACITY
                var capacityId = _auth.GetCapacityId(); // I'll add this helper to Auth service
                if (!string.IsNullOrEmpty(capacityId))
                {
                    Console.WriteLine($"[SERVICE] Auto-assigning to capacity: {capacityId}");
                    try {
                        await client.Groups.AssignToCapacityAsync(newGroup.Id, new AssignToCapacityRequest { CapacityId = Guid.Parse(capacityId) });
                        Console.WriteLine("[SERVICE] Successfully assigned to capacity.");
                    } catch (Exception ex) {
                        Console.WriteLine($"[SERVICE] Capacity assignment failed: {ex.Message}. You may need to assign it manually.");
                    }
                }

                // 🔑 CRITICAL: Add the human admin user to the workspace
                // Since SP is the creator, the human user won't see it otherwise.
                var adminEmail = _auth.GetAdminEmail();
                if (!string.IsNullOrEmpty(adminEmail))
                {
                    try 
                    {
                        Console.WriteLine($"[SERVICE] Adding {adminEmail} as Admin to new workspace...");
                        await client.Groups.AddGroupUserAsync(newGroup.Id, new GroupUser 
                        { 
                            Identifier = adminEmail, 
                            GroupUserAccessRight = "Admin", 
                            PrincipalType = "User" 
                        });
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[SERVICE] Warning: Could not add {adminEmail} as admin: {ex.Message}");
                    }
                }

                return newGroup;
            }
            catch (HttpOperationException ex)
            {
                Console.WriteLine($"[SERVICE] ERROR in GetOrCreateWorkspace: {ex.Response.Content}");
                throw;
            }
        }

        // 🔄 Sync Workspaces with Local DB
        public async Task SyncWorkspaces(int userId, AppDbContext db)
        {
            try 
            {
                Console.WriteLine($"[SERVICE] Syncing workspaces for User ID: {userId}");
                var client = await GetClient();
                var pbiGroups = await client.Groups.GetGroupsAsync();
                var pbiWorkspaces = pbiGroups.Value;

                var localWorkspaces = db.Workspaces.Where(w => w.UserId == userId).ToList();

                // 1. Add/Update from Power BI to Local DB
                var adminEmail = _auth.GetAdminEmail();
                foreach (var pbiWs in pbiWorkspaces)
                {
                    // Attempt to ensure AdminEmail is always an admin (Visibility Fix)
                    if (!string.IsNullOrEmpty(adminEmail))
                    {
                        try 
                        {
                            await client.Groups.AddGroupUserAsync(pbiWs.Id, new GroupUser 
                            { 
                                Identifier = adminEmail, 
                                GroupUserAccessRight = "Admin", 
                                PrincipalType = "User" 
                            });
                        }
                        catch { /* Ignore if already admin or error */ }
                    }

                    var existing = localWorkspaces.FirstOrDefault(w => w.PowerBIWorkspaceId == pbiWs.Id.ToString());
                    if (existing == null)
                    {
                        Console.WriteLine($"[SERVICE] Adding new PBI workspace to DB: {pbiWs.Name}");
                        db.Workspaces.Add(new Workspace
                        {
                            Name = pbiWs.Name,
                            PowerBIWorkspaceId = pbiWs.Id.ToString(),
                            UserId = userId
                        });
                    }
                    else if (existing.Name != pbiWs.Name)
                    {
                        Console.WriteLine($"[SERVICE] Updating workspace name in DB: {existing.Name} -> {pbiWs.Name}");
                        existing.Name = pbiWs.Name;
                    }
                }

                // 2. Remove from Local DB if deleted in Power BI
                foreach (var localWs in localWorkspaces)
                {
                    if (!pbiWorkspaces.Any(pbi => pbi.Id.ToString() == localWs.PowerBIWorkspaceId))
                    {
                        Console.WriteLine($"[SERVICE] Removing workspace from DB (deleted in PBI): {localWs.Name}");
                        db.Workspaces.Remove(localWs);
                    }
                }

                await db.SaveChangesAsync();
            }
            catch (HttpOperationException ex)
            {
                Console.WriteLine($"[SERVICE] Power BI API Error during Sync: {ex.Response.Content}");
                throw new Exception($"Power BI API Error: {ex.Response.ReasonPhrase}. Details: {ex.Response.Content}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SERVICE] General Error during Sync: {ex.Message}");
                throw;
            }
        }

        // ✏️ Rename Workspace in Power BI
        public async Task RenameWorkspace(Guid workspaceId, string newName)
        {
            Console.WriteLine($"[SERVICE] Renaming workspace {workspaceId} to '{newName}'");
            var client = await GetClient();
            // Power BI API doesn't have a direct 'Rename' method in the SDK for Groups, 
            // but you can update group properties if it's a modern workspace.
            // For simplicity in this demo, we assume modern workspaces.
            // Note: UpdateGroup is often restricted for Service Principals depending on tenant settings.
            // We use the REST API via the client.
            await client.Groups.UpdateGroupAsync(workspaceId, new UpdateGroupRequest { Name = newName });
        }

        // 🗑️ Delete Workspace from Power BI
        public async Task DeleteWorkspace(Guid workspaceId)
        {
            Console.WriteLine($"[SERVICE] Deleting workspace {workspaceId}");
            var client = await GetClient();
            await client.Groups.DeleteGroupAsync(workspaceId);
        }

        // 🔄 Sync Reports within a Workspace
        public async Task SyncReports(int localWorkspaceId, Guid pbiWorkspaceId, AppDbContext db)
        {
            try 
            {
                Console.WriteLine($"[SERVICE] Syncing reports for Workspace: {pbiWorkspaceId}");
                var client = await GetClient();
                var pbiReports = (await client.Reports.GetReportsInGroupAsync(pbiWorkspaceId)).Value;

                var localReports = db.Reports.Where(r => r.WorkspaceId == localWorkspaceId).ToList();

                // 1. Add/Update from Power BI to Local DB
                foreach (var pbiRep in pbiReports)
                {
                    var existing = localReports.FirstOrDefault(r => r.PowerBIReportId == pbiRep.Id.ToString());
                    if (existing == null)
                    {
                        Console.WriteLine($"[SERVICE] Adding new PBI report to DB: {pbiRep.Name} (Type: {pbiRep.ReportType})");
                        db.Reports.Add(new PowerBI.Models.Report
                        {
                            Name = pbiRep.Name,
                            PowerBIReportId = pbiRep.Id.ToString(),
                            PowerBIDatasetId = pbiRep.DatasetId,
                            WorkspaceId = localWorkspaceId,
                            ReportType = pbiRep.ReportType == "PaginatedReport" ? "RDL" : "PowerBI"
                        });
                    }
                    else
                    {
                        if (existing.Name != pbiRep.Name) existing.Name = pbiRep.Name;
                        var correctType = pbiRep.ReportType == "PaginatedReport" ? "RDL" : "PowerBI";
                        if (existing.ReportType != correctType) existing.ReportType = correctType;
                    }
                }

                // 2. Remove from Local DB if deleted in Power BI
                foreach (var localRep in localReports)
                {
                    if (!pbiReports.Any(pbi => pbi.Id.ToString() == localRep.PowerBIReportId))
                    {
                        Console.WriteLine($"[SERVICE] Removing report from DB (deleted in PBI): {localRep.Name}");
                        db.Reports.Remove(localRep);
                    }
                }

                await db.SaveChangesAsync();
            }
            catch (HttpOperationException ex)
            {
                Console.WriteLine($"[SERVICE] Power BI API Error during SyncReports: {ex.Response.Content}");
                throw new Exception($"Power BI API Error: {ex.Response.ReasonPhrase}. Details: {ex.Response.Content}");
            }
        }

        public async Task<Import> UploadReport(Guid workspaceId, string name, Stream stream)
        {
            try 
            {
                Console.WriteLine($"[SERVICE] Importing {name} to Workspace {workspaceId}");
                var client = await GetClient();

                var isRdl = name.EndsWith(".rdl", StringComparison.OrdinalIgnoreCase);
                var datasetName = isRdl ? name : Path.GetFileNameWithoutExtension(name);

                Console.WriteLine($"[SERVICE] Stream length: {stream.Length} bytes");

                if (stream.Length == 0)
                    throw new Exception("The uploaded file is empty.");

                // For RDL files via Service Principal, 'Abort' is the safest mode.
                // IMPORTANT: RDL files often REQUIRE the .rdl extension in the display name to be recognized correctly.
                var conflictMode = isRdl ? ImportConflictHandlerMode.Abort : ImportConflictHandlerMode.CreateOrOverwrite;
                var finalDisplayName = datasetName; // Keep the name as-is (including .rdl for Paginated)

                Console.WriteLine($"[SERVICE] Uploading {name} as {finalDisplayName} (Mode: {conflictMode})");

                var import = await client.Imports.PostImportWithFileAsync(
                    workspaceId,
                    stream,
                    datasetDisplayName: finalDisplayName,
                    nameConflict: conflictMode);

                return import;
            }
            catch (HttpOperationException ex)
            {
                Console.WriteLine($"[SERVICE] Power BI API Error during Upload: {ex.Response.Content}");
                var details = ex.Response.Content;
                
                if (details.Contains("ImportUnsupportedOptionError"))
                {
                    throw new Exception("Power BI Error: 'ImportUnsupportedOptionError'. This usually means the workspace is NOT on a Fabric/Premium Capacity. Please ensure the workspace is assigned to your Fabric Trial capacity in the Power BI Portal.");
                }

                throw new Exception($"Power BI Upload Error: {ex.Response.ReasonPhrase}. Details: {details}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Upload failed: {ex.Message}");
            }
        }

        // 📤 Export PDF to Stream
        public async Task<Stream> ExportReportAsStream(Guid workspaceId, Guid reportId)
        {
            Console.WriteLine($"[SERVICE] Exporting Report {reportId} to PDF...");
            var client = await GetClient();

            var exportRequest = new ExportReportRequest { Format = FileFormat.PDF };
            var export = await client.Reports.ExportToFileInGroupAsync(workspaceId, reportId, exportRequest);

            Export status;
            do
            {
                await Task.Delay(3000);
                status = await client.Reports.GetExportToFileStatusInGroupAsync(workspaceId, reportId, export.Id);
                Console.WriteLine($"[SERVICE] Export Status: {status.Status}");
            } while (status.Status != ExportState.Succeeded && status.Status != ExportState.Failed);

            if (status.Status == ExportState.Failed) throw new Exception("Power BI Export Failed.");

            return await client.Reports.GetFileOfExportToFileInGroupAsync(workspaceId, reportId, export.Id);
        }

        public async Task<ReportEmbedConfig> GetEmbedConfig(Guid workspaceId, Guid reportId)
        {
            Console.WriteLine($"[SERVICE] Fetching Embed Config for Report {reportId}");
            var client = await GetClient();

            var report = await client.Reports.GetReportInGroupAsync(workspaceId, reportId);
            
            var generateTokenRequestParameters = new GenerateTokenRequest(accessLevel: "view");
            var tokenResponse = await client.Reports.GenerateTokenInGroupAsync(workspaceId, reportId, generateTokenRequestParameters);

            Console.WriteLine("[SERVICE] Embed Token generated successfully.");

            return new ReportEmbedConfig
            {
                ReportId = report.Id.ToString(),
                DatasetId = report.DatasetId,
                EmbedUrl = report.EmbedUrl,
                EmbedToken = tokenResponse.Token,
                ReportName = report.Name
            };
        }


        public async Task DiscoverReportFilters(int reportId, Guid? datasetId, AppDbContext db)
        {
            var report = await db.Reports.FindAsync(reportId);
            if (report == null) throw new Exception("Report not found in local DB.");

            var workspace = await db.Workspaces.FindAsync(report.WorkspaceId);
            if (workspace == null || string.IsNullOrEmpty(workspace.PowerBIWorkspaceId))
                throw new Exception("Workspace not found or not synced.");

            Console.WriteLine($"\n[SCHEMA-DISCOVERY] ==============================");
            Console.WriteLine($"[SCHEMA-DISCOVERY] MODE: {report.ReportType}");
            Console.WriteLine($"[SCHEMA-DISCOVERY] Report Name: {report.Name}");
            Console.WriteLine($"[SCHEMA-DISCOVERY] ==============================");

            // 1. Clear stale DB data
            var existing = db.ReportFilters.Where(f => f.ReportId == reportId).ToList();
            Console.WriteLine($"[SCHEMA-DISCOVERY] Purging {existing.Count} stale records for reportId={reportId}");
            db.ReportFilters.RemoveRange(existing);
            await db.SaveChangesAsync();

            int count = 0;

            if (report.ReportType == "RDL")
            {
                // DISCOVERY FOR PAGINATED REPORTS (RDL)
                // First try the API, if it fails, fallback to parsing the local XML file
                try
                {
                    Console.WriteLine($"[SCHEMA-DISCOVERY] RDL: Attempting API discovery...");
                    var token = await _auth.GetAccessToken();
                    using var httpClient = new HttpClient();
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    var url = $"https://api.powerbi.com/v1.0/myorg/groups/{workspace.PowerBIWorkspaceId}/reports/{report.PowerBIReportId}/parameters";

                    var response = await httpClient.GetAsync(url);
                    if (response.IsSuccessStatusCode)
                    {
                        var rawJson = await response.Content.ReadAsStringAsync();
                        var root = Newtonsoft.Json.Linq.JObject.Parse(rawJson);
                        var parameters = root["value"] as Newtonsoft.Json.Linq.JArray;

                        if (parameters != null)
                        {
                            foreach (var param in parameters)
                            {
                                string? paramName = param["name"]?.ToString();
                                if (string.IsNullOrEmpty(paramName)) continue;

                                db.ReportFilters.Add(new ReportFilter
                                {
                                    ReportId = reportId,
                                    TableName = "RDL_PARAMETER",
                                    ColumnName = paramName,
                                    DisplayName = paramName,
                                    IsActive = true
                                });
                                count++;
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"[SCHEMA-DISCOVERY] RDL: API returned {response.StatusCode}. Falling back to Local XML Parsing...");
                        // FALLBACK: Parse the local .rdl file
                        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", $"{report.Name}.rdl");
                        if (File.Exists(filePath))
                        {
                            var xml = System.Xml.Linq.XDocument.Load(filePath);
                            var ns = xml.Root?.GetDefaultNamespace();
                            var rdlParams = xml.Descendants(ns + "ReportParameter");

                            foreach (var p in rdlParams)
                            {
                                string? paramName = p.Attribute("Name")?.Value;
                                if (string.IsNullOrEmpty(paramName)) continue;

                                db.ReportFilters.Add(new ReportFilter
                                {
                                    ReportId = reportId,
                                    TableName = "RDL_PARAMETER",
                                    ColumnName = paramName,
                                    DisplayName = paramName,
                                    IsActive = true
                                });
                                count++;
                            }
                            Console.WriteLine($"[SCHEMA-DISCOVERY] RDL: Local parsing found {count} parameters.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[SCHEMA-DISCOVERY] RDL Error: {ex.Message}");
                    throw new Exception($"Failed to discover RDL parameters: {ex.Message}");
                }
            }
            else
            {
                // DISCOVERY FOR POWER BI REPORTS - Use Datasets Tables API
                if (!datasetId.HasValue) throw new Exception("Dataset ID is missing for Power BI report.");

                string rawJson;
                try
                {
                    var token = await _auth.GetAccessToken();
                    using var httpClient = new HttpClient();
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    var url = $"https://api.powerbi.com/v1.0/myorg/datasets/{datasetId}/tables";

                    Console.WriteLine($"[SCHEMA-DISCOVERY] Calling: GET {url}");
                    var response = await httpClient.GetAsync(url);
                    rawJson = await response.Content.ReadAsStringAsync();

                    if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                    {
                        throw new Exception("Access Denied (403). For Service Principals, this usually means the workspace is NOT on a Premium/Fabric capacity, or the 'XMLA Endpoint' setting is not 'Read Write'. Please wait a few minutes after assigning the capacity and try again.");
                    }

                    if (!response.IsSuccessStatusCode)
                        throw new Exception($"Power BI REST API returned {(int)response.StatusCode}: {rawJson}");

                    var root = Newtonsoft.Json.Linq.JObject.Parse(rawJson);
                    var tables = root["value"] as Newtonsoft.Json.Linq.JArray;

                    if (tables != null)
                    {
                        foreach (var tbl in tables)
                        {
                            string? tableName = tbl["name"]?.ToString();
                            if (string.IsNullOrEmpty(tableName)) continue;

                            var cols = tbl["columns"] as Newtonsoft.Json.Linq.JArray;
                            if (cols != null)
                            {
                                foreach (var col in cols)
                                {
                                    string? colName = col["name"]?.ToString();
                                    if (string.IsNullOrEmpty(colName)) continue;

                                    db.ReportFilters.Add(new ReportFilter
                                    {
                                        ReportId = reportId,
                                        TableName = tableName,
                                        ColumnName = colName,
                                        DisplayName = colName,
                                        IsActive = count < 6
                                    });
                                    count++;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Failed to discover Power BI schema: {ex.Message}");
                }
            }

            if (count == 0) throw new Exception("No queryable fields or parameters discovered.");

            await db.SaveChangesAsync();
            Console.WriteLine($"[SCHEMA-DISCOVERY] SUCCESS: {count} field(s) persisted to DB.");
        }




        public async Task<List<string>> GetColumnValues(Guid? datasetId, string tableName, string columnName, Guid? reportId = null, Guid? workspaceId = null)
        {
            if (string.IsNullOrEmpty(tableName) || string.IsNullOrEmpty(columnName))
                throw new Exception("[DYNAMIC] Error: Missing Table or Column name.");

            if (tableName == "RDL_PARAMETER")
            {
                // Logic for Paginated Report Parameters
                if (!reportId.HasValue || !workspaceId.HasValue) throw new Exception("Report/Workspace ID required for RDL values.");
                
                try
                {
                    var token = await _auth.GetAccessToken();
                    using var httpClient = new HttpClient();
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    var url = $"https://api.powerbi.com/v1.0/myorg/groups/{workspaceId}/reports/{reportId}/parameters";

                    var response = await httpClient.GetAsync(url);
                    var rawJson = await response.Content.ReadAsStringAsync();
                    var root = Newtonsoft.Json.Linq.JObject.Parse(rawJson);
                    var parameters = root["value"] as Newtonsoft.Json.Linq.JArray;

                    if (parameters != null)
                    {
                        var param = parameters.FirstOrDefault(p => p["name"]?.ToString() == columnName);
                        if (param != null)
                        {
                            var suggested = param["suggestedValues"] as Newtonsoft.Json.Linq.JArray;
                            if (suggested != null)
                            {
                                return suggested.Select(v => v.ToString()).ToList();
                            }
                        }
                    }
                    return new List<string>();
                }
                catch { return new List<string>(); }
            }

            // Standard DAX Logic for Power BI Reports
            if (!datasetId.HasValue) throw new Exception("Dataset ID required for Power BI values.");

            var daxQuery = $"EVALUATE DISTINCT('{tableName}'[{columnName}])";
            
            Console.WriteLine("--------------------------------------------------");
            Console.WriteLine($"[DYNAMIC] PROBE: '{tableName}'[{columnName}]");
            Console.WriteLine($"[DYNAMIC] DAX: {daxQuery}");

            var results = new List<string>();
            try
            {
                var client = await GetClient();
                var request = new DatasetExecuteQueriesRequest(new List<DatasetExecuteQueriesQuery> { new DatasetExecuteQueriesQuery(daxQuery) });
                var response = await client.Datasets.ExecuteQueriesAsync(datasetId.ToString(), request);

                if (response?.Results != null && response.Results.Count > 0)
                {
                    var table = response.Results[0].Tables[0];
                    foreach (dynamic row in table.Rows)
                    {
                        var rowStr = row?.ToString();
                        if (string.IsNullOrEmpty(rowStr)) continue;

                        var dict = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(rowStr);
                        if (dict != null && dict.Values.Count > 0)
                        {
                            foreach (var valObj in dict.Values)
                            {
                                string? valStr = valObj?.ToString();
                                if (!string.IsNullOrEmpty(valStr))
                                {
                                    results.Add(valStr);
                                    break; 
                                }
                            }
                        }
                    }
                }
                Console.WriteLine($"[DYNAMIC] RESULT: {results.Count} values found.");
                Console.WriteLine("--------------------------------------------------");
            }
            catch (Exception ex) 
            { 
                if (ex.Message.Contains("403") || ex.Message.Contains("Forbidden"))
                {
                    Console.WriteLine("[DYNAMIC] 403 ERROR: Service Principal requires Premium/Fabric capacity to access dataset metadata/queries.");
                }
                Console.WriteLine($"[DYNAMIC] DAX FAIL: {ex.Message}"); 
            }
            return results;
        }


    }
}