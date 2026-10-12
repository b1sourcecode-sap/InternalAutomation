using ItemCreationInterCompanyUtility;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Sap.Data.Hana;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace IRDSync
{
    class IRDSyncDairy
    {


        //public void StartProcess()
        //{
        //    string excelPath = ConfigurationManager.AppSettings["ConfigExcelPath"];
        //    List<TenantConfig> tenants = LoadTenantsFromExcel(excelPath);

        //    foreach (var tenant in tenants)
        //    {
        //        Console.WriteLine($"--- Processing: {tenant.CompanyDB} on {tenant.Server} ---");
        //        RunSyncForTenant(tenant);
        //    }
        //}

        //private List<TenantConfig> LoadTenantsFromExcel(string path)
        //{
        //    List<TenantConfig> list = new List<TenantConfig>();
        //    Application xlApp = new Application();
        //    Workbook xlWorkbook = xlApp.Workbooks.Open(path);
        //    _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        //    Range xlRange = xlWorksheet.UsedRange;

        //    // Assuming: Col 1:Server, 2:DBUser, 3:DBPass, 4:DBName, 5:IRDUser, 6:IRDPass, 7:PAN, 8:Fiscal
        //    for (int i = 2; i <= xlRange.Rows.Count; i++)
        //    {
        //        if (xlRange.Cells[i, 1].Value2 == null) break;
        //        list.Add(new TenantConfig
        //        {
        //            Server = xlRange.Cells[i, 1].Value2.ToString(),
        //            DBUser = xlRange.Cells[i, 2].Value2.ToString(),
        //            DBPassword = xlRange.Cells[i, 3].Value2.ToString(),
        //            CompanyDB = xlRange.Cells[i, 4].Value2.ToString(),
        //            IRDUser = xlRange.Cells[i, 5].Value2.ToString(),
        //            IRDPassword = xlRange.Cells[i, 6].Value2.ToString(),
        //            SellerPan = xlRange.Cells[i, 7].Value2.ToString(),
        //            FiscalYear = xlRange.Cells[i, 8].Value2.ToString()
        //        });
        //    }
        //    xlWorkbook.Close();
        //    xlApp.Quit();
        //    return list;
        //}

        //private void RunSyncForTenant(TenantConfig tenant)
        //{
        //    string connStr = $"Server={tenant.Server};UserID={tenant.DBUser};Password={tenant.DBPassword};CurrentSchema={tenant.CompanyDB}";
        //    string today = DateTime.Now.ToString("yyyyMMdd");

        //    string query = $@"SELECT top 1 ""DocEntry"" FROM OINV WHERE ""DocDate"" = '{today}' AND ""U_IsSync"" = 'Y' AND ""U_SyncIRD"" = 'N'";
        //    DataSet ds = SqlHelper.Hana_ADOConnection(query, connStr);

        //    if (ds != null && ds.Tables[0].Rows.Count > 0)
        //    {
        //        foreach (DataRow row in ds.Tables[0].Rows)
        //        {
        //            SyncInvoice(row["DocEntry"].ToString(), tenant, connStr);
        //        }
        //    }
        //}

        //private void SyncInvoice(string docEntry, TenantConfig tenant, string connStr)
        //{
        //    try
        //    {
        //        string sql = $@"select Top 1 OINV.""CardName"", OINV.""DocTotal"", OINV.""VatSum"", 
        //                      (OINV.""DocTotal""-OINV.""VatSum"") as ""AmtBeforeTax"", 
        //                      OINV.""DocNum"", OINV.""U_LD"", OCRD.""LicTradNum"" 
        //                      from OINV inner join OCRD ON OCRD.""CardCode"" = OINV.""CardCode"" 
        //                      INNER JOIN NNM1 ON OINV.""Series"" = NNM1.""Series"" AND OINV.""ObjType"" = NNM1.""ObjectCode"" 
        //                      where OINV.""DocEntry"" = '{docEntry}'";

        //        DataSet ds = SqlHelper.Hana_ADOConnection(sql, connStr);
        //        if (ds == null || ds.Tables[0].Rows.Count == 0) return;
        //        DataRow r = ds.Tables[0].Rows[0];

        //        var payload = new
        //        {
        //            username = tenant.IRDUser,
        //            password = tenant.IRDPassword,
        //            seller_pan = tenant.SellerPan,
        //            buyer_pan = r["LicTradNum"].ToString().Trim(),
        //            buyer_name = r["CardName"].ToString().Trim(),
        //            fiscal_year = tenant.FiscalYear,
        //            invoice_number = r["DocNum"].ToString().Trim(),
        //            invoice_date = r["U_LD"].ToString().Trim(),
        //            total_sales = Convert.ToDouble(r["DocTotal"]),
        //            taxable_sales_vat = Convert.ToDouble(r["AmtBeforeTax"]),
        //            vat = Convert.ToDouble(r["VatSum"]),
        //            isrealtime = true,
        //            datetimeClient = DateTime.Now
        //        };

        //        using (var client = new HttpClient())
        //        {
        //            // Global API settings from App.config
        //            client.BaseAddress = new Uri(ConfigurationManager.AppSettings["InvBaseAddress"]);
        //            string json = JsonConvert.SerializeObject(payload);
        //            var content = new StringContent(json, Encoding.UTF8, "application/json");

        //            var response = client.PostAsync(ConfigurationManager.AppSettings["Invresponse"], content).Result;
        //            string resBody = response.Content.ReadAsStringAsync().Result;

        //            string status = (response.IsSuccessStatusCode && resBody.Contains("200")) ? "Y" : "N";
        //            string msg = (status == "Y") ? "200" : resBody.Replace("'", "");

        //            // Update HANA back
        //            using (HanaConnection conn = new HanaConnection(connStr))
        //            {
        //                conn.Open();
        //                string update = $@"UPDATE OINV SET ""U_SyncIRD""='{status}', ""U_SyncMsgID""='{msg}' WHERE ""DocEntry""='{docEntry}'";
        //                new HanaCommand(update, conn).ExecuteNonQuery();
        //            }
        //        }
        //    }
        //    catch (Exception ex) { Console.WriteLine("Sync Error: " + ex.Message); }
        //}

        public void StartProcess()
        {
            Logger.Log("===== GLOBAL PROCESS START =====", "General");
            try
            {
                string excelPath = ConfigurationManager.AppSettings["ConfigExcelPath"];
                Logger.Log($"Reading configuration from: {excelPath}", "General");

                List<TenantConfig> tenants = LoadTenantsFromExcel(excelPath);
                Logger.Log($"Successfully loaded {tenants.Count} databases from Excel.", "General");

                foreach (var tenant in tenants)
                {
                    Logger.Log($"--- STARTING SYNC FOR DB: {tenant.CompanyDB} ---", tenant.CompanyDB);
                    RunSyncForTenant(tenant);
                    Logger.Log($"--- FINISHED SYNC FOR DB: {tenant.CompanyDB} ---", tenant.CompanyDB);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Critical error in StartProcess", "General", ex);
            }
            Logger.Log("===== GLOBAL PROCESS FINISHED =====", "General");
        }

        private List<TenantConfig> LoadTenantsFromExcel(string path)
        {
            List<TenantConfig> list = new List<TenantConfig>();
            try
            {
                using (var workbook = new XLWorkbook(path))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RangeUsed().RowsUsed();

                    foreach (var row in rows.Skip(1)) // Skip header row
                    {
                        // Verify the first cell isn't empty
                        if (row.Cell(1).IsEmpty()) break;

                        list.Add(new TenantConfig
                        {
                            Server = row.Cell(1).GetValue<string>().Trim(),
                            DBUser = row.Cell(2).GetValue<string>().Trim(),
                            DBPassword = row.Cell(3).GetValue<string>().Trim(),
                            CompanyDB = row.Cell(4).GetValue<string>().Trim(),
                            IRDUser = row.Cell(5).GetValue<string>().Trim(),
                            IRDPassword = row.Cell(6).GetValue<string>().Trim(),
                            SellerPan = row.Cell(7).GetValue<string>().Trim(),
                            FiscalYear = row.Cell(8).GetValue<string>().Trim()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to read Excel configuration file using ClosedXML.", "General", ex);
            }
            return list;
        }

        private void RunSyncForTenant(TenantConfig tenant)
        {
            // Note: CurrentSchema (No space) is the correct property for Hana ADO provider
            string connStr = $"Server={tenant.Server};UserID={tenant.DBUser};Password={tenant.DBPassword};CurrentSchema={tenant.CompanyDB}";
            string today = DateTime.Now.ToString("yyyyMMdd");

            try
            {
                string query = "SELECT \"DocEntry\" FROM OINV WHERE \"DocDate\" >= TO_DATE(LEFT(CURRENT_DATE, 7) || '-01')";
                DataSet ds = ExecuteHanaQuery(query, connStr, tenant.CompanyDB);

                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    Logger.Log($"Found {ds.Tables[0].Rows.Count} pending invoices.", tenant.CompanyDB);
                    foreach (DataRow row in ds.Tables[0].Rows)
                    {
                        SyncInvoice(row["DocEntry"].ToString(), tenant, connStr);
                    }
                }
                else
                {
                    Logger.Log("No pending invoices found for today.", tenant.CompanyDB);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Database Connection Failure for {tenant.CompanyDB}. Check credentials.", tenant.CompanyDB, ex);
            }
        }

        private void SyncInvoice(string docEntry, TenantConfig tenant, string connStr)
        {
            try
            {
                Logger.Log($"Processing Invoice DocEntry: {docEntry}", tenant.CompanyDB);

                string sql = $@"SELECT Top 1 OINV.""CardName"", OINV.""DocTotal"", OINV.""VatSum"", 
                              (OINV.""DocTotal""-OINV.""VatSum"") as ""AmtBeforeTax"", 
                              OINV.""DocNum"", OINV.""U_LD"", OCRD.""LicTradNum"" 
                              FROM OINV INNER JOIN OCRD ON OCRD.""CardCode"" = OINV.""CardCode"" 
                              INNER JOIN NNM1 ON OINV.""Series"" = NNM1.""Series"" AND OINV.""ObjType"" = NNM1.""ObjectCode"" 
                              WHERE OINV.""DocEntry"" = '{docEntry}'";

                DataSet ds = ExecuteHanaQuery(sql, connStr, tenant.CompanyDB);
                if (ds == null || ds.Tables[0].Rows.Count == 0) return;

                DataRow r = ds.Tables[0].Rows[0];

                var payload = new
                {
                    username = tenant.IRDUser,
                    password = tenant.IRDPassword,
                    seller_pan = tenant.SellerPan,
                    buyer_pan = r["LicTradNum"].ToString().Trim(),
                    buyer_name = r["CardName"].ToString().Trim(),
                    fiscal_year = tenant.FiscalYear,
                    invoice_number = r["DocNum"].ToString().Trim(),
                    invoice_date = r["U_LD"].ToString().Trim(),
                    total_sales = Convert.ToDouble(r["DocTotal"]),
                    taxable_sales_vat = Convert.ToDouble(r["AmtBeforeTax"]),
                    vat = Convert.ToDouble(r["VatSum"]),
                    isrealtime = true,
                    datetimeClient = DateTime.Now
                };

                using (var client = new HttpClient())
                {
                    string baseUri = ConfigurationManager.AppSettings["InvBaseAddress"];
                    string endpoint = ConfigurationManager.AppSettings["Invresponse"];
                    client.BaseAddress = new Uri(baseUri);

                    string json = JsonConvert.SerializeObject(payload);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = client.PostAsync(endpoint, content).Result;
                    string resBody = response.Content.ReadAsStringAsync().Result;

                    string status = "N";
                    string msg = resBody.Replace("'", ""); // Default to error message

                    // Check logic for 200 and 101
                    if (response.IsSuccessStatusCode)
                    {
                        if (resBody.Contains("200"))
                        {
                            status = "Y";
                            msg = "200";
                        }
                        else if (resBody.Contains("101"))
                        {
                            status = "Y";
                            msg = "101";
                        }
                    }

                    Logger.Log($"API Result for {docEntry}: {status} - {msg}", tenant.CompanyDB);

                    // Update SAP HANA
                    using (HanaConnection conn = new HanaConnection(connStr))
                    {
                        conn.Open();
                        // Using parameters is highly recommended to prevent SQL errors
                        string update = @"UPDATE OINV SET ""U_SyncIRD"" = ?, ""U_SyncMsgID"" = ? WHERE ""DocEntry"" = ?";

                        using (HanaCommand cmd = new HanaCommand(update, conn))
                        {
                            cmd.Parameters.Add(new HanaParameter("", status));
                            cmd.Parameters.Add(new HanaParameter("", msg));
                            cmd.Parameters.Add(new HanaParameter("", docEntry));

                            cmd.ExecuteNonQuery();
                        }
                        Logger.Log($"HANA status updated for {docEntry}", tenant.CompanyDB);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Sync Error for DocEntry {docEntry}", tenant.CompanyDB, ex);
            }
        }

        // Internal helper to avoid code duplication
        private DataSet ExecuteHanaQuery(string query, string connStr, string dbName)
        {
            try
            {
                using (HanaConnection conn = new HanaConnection(connStr))
                {
                    conn.Open();
                    using (HanaDataAdapter da = new HanaDataAdapter(query, conn))
                    {
                        DataSet ds = new DataSet();
                        da.Fill(ds);
                        return ds;
                    }
                }
            }
            catch (TypeInitializationException tex)
            {
                // THIS IS THE IMPORTANT PART
                string inner = tex.InnerException != null ? tex.InnerException.Message : "No inner exception";
                Logger.Error($"CRITICAL DRIVER ERROR: {inner}", dbName);
                throw;
            }
            catch (Exception ex)
            {
                Logger.Error($"HANA Query Error on {dbName}", dbName, ex);
                throw; // Rethrow to let RunSyncForTenant catch it
            }
        }
    }


    //public static class Logger
    //{
    //    private static string baseLogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

    //    // dbName is optional. If empty, it goes to a "General" folder.
    //    public static void Log(string message, string dbName = "General", string level = "INFO")
    //    {
    //        try
    //        {
    //            // Create path: Logs/DatabaseName/
    //            string targetFolder = Path.Combine(baseLogPath, dbName);
    //            if (!Directory.Exists(targetFolder)) Directory.CreateDirectory(targetFolder);

    //            string fileName = $"SyncLog_{DateTime.Now:yyyyMMdd}.txt";
    //            string fullPath = Path.Combine(targetFolder, fileName);

    //            string logLine = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}";

    //            Console.WriteLine($"[{dbName}] {logLine}");
    //            File.AppendAllText(fullPath, logLine + Environment.NewLine);
    //        }
    //        catch { /* Prevent logger from crashing the app */ }
    //    }

    //    public static void Error(string message, string dbName = "General", Exception ex = null)
    //    {
    //        string detail = ex != null ? $"\nException: {ex.Message}\nStacktrace: {ex.StackTrace}" : "";
    //        Log($"{message} {detail}", dbName, "ERROR");
    //    }
    //}

    public static class Logger
    {
        private static string GetLogPath()
        {
            string path = ConfigurationManager.AppSettings["LogPath"];
            // Fallback to local folder if config is missing
            if (string.IsNullOrEmpty(path))
                path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
            return path;
        }

        public static void Log(string message, string dbName = "General", string level = "INFO")
        {
            try
            {
                string baseLogPath = GetLogPath();
                string targetFolder = Path.Combine(baseLogPath, dbName);

                if (!Directory.Exists(targetFolder))
                    Directory.CreateDirectory(targetFolder);

                string fileName = $"SyncLog_{DateTime.Now:yyyyMMdd}.txt";
                string fullPath = Path.Combine(targetFolder, fileName);

                string logLine = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}";

                Console.WriteLine($"[{dbName}] {logLine}");
                File.AppendAllText(fullPath, logLine + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Log Error: " + ex.Message);
            }
        }

        public static void Error(string message, string dbName = "General", Exception ex = null)
        {
            string detail = ex != null ? $"\nException: {ex.Message}\nStacktrace: {ex.StackTrace}" : "";
            Log($"{message} {detail}", dbName, "ERROR");
        }
    }


}
