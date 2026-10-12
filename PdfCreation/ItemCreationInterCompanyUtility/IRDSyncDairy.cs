using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;

namespace PDFCreation
{
    class PDFCreationDairy
    {
        // One row from the Excel schedule
        private class ReportDefinition
        {
            public string Server;
            public string Database;
            public string User;
            public string Password;
            public string RptPath;
            public string ToEmail;
            public string CcEmail;
            public TimeSpan ScheduledTime;   // parsed from "HH:mm"

            public string Subject;  
            public string Body;
            public string PdfName;

            // Used for log folder + pdf filename
            public string Key => string.IsNullOrEmpty(Database) ? "Report" : Database;
        }

        public void StartProcess()
        {
            Logger.Log("===== FLASH REPORT SCHEDULER START =====", "General");

            string excelPath = ConfigurationManager.AppSettings["ScheduleExcelPath"];
            List<ReportDefinition> reports = LoadReportsFromExcel(excelPath);
            Logger.Log($"Loaded {reports.Count} report row(s) from Excel.", "General");

            TimeSpan now = DateTime.Now.TimeOfDay;

            // Exact match on hour + minute
            var due = reports
                .Where(r => r.ScheduledTime.Hours == now.Hours
                         && r.ScheduledTime.Minutes == now.Minutes)
                .ToList();

            if (due.Count == 0)
            {
                Logger.Log($"No reports scheduled for {DateTime.Now:HH:mm}.", "General");
                Logger.Log("===== FLASH REPORT SCHEDULER FINISHED =====", "General");
                return;
            }

            bool allSucceeded = true;
            foreach (var report in due)
            {
                Logger.Log($"{report.Key}: scheduled {report.ScheduledTime:hh\\:mm} matches now, running.", report.Key);
                bool ok = RunReport(report);
                if (!ok) allSucceeded = false;
            }

            // Completion email only if everything that ran succeeded
            if (allSucceeded)
            {
                try
                {
                    SendCompletionEmail(due);
                    Logger.Log("Completion notification email sent.", "General");
                }
                catch (Exception ex)
                {
                    Logger.Error("Failed to send completion notification email.", "General", ex);
                }
            }
            else
            {
                Logger.Log("One or more reports failed; completion email NOT sent.", "General");
            }

            Logger.Log("===== FLASH REPORT SCHEDULER FINISHED =====", "General");
        }

        // ---- Excel loading ----
        private List<ReportDefinition> LoadReportsFromExcel(string path)
        {
            var list = new List<ReportDefinition>();
            try
            {
                if (string.IsNullOrEmpty(path) || !File.Exists(path))
                {
                    Logger.Error($"Schedule Excel not found: {path}", "General");
                    return list;
                }

                using (var workbook = new XLWorkbook(path))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RangeUsed().RowsUsed();

                    // Columns: 1 Server | 2 DatabaseName | 3 UserName | 4 Password
                    //          5 Report(.rpt path) | 6 ToEmail | 7 CCEmail | 8 Time (HH:mm)
                    foreach (var row in rows.Skip(1)) // skip header
                    {
                        if (row.Cell(1).IsEmpty()) break;

                        string timeText = row.Cell(8).GetValue<string>().Trim();
                        TimeSpan scheduled;
                        if (!TryParseTime(row.Cell(8), out scheduled))
                        {
                            Logger.Error(
                                $"Skipping row: bad Time value '{timeText}' for DB '{row.Cell(2).GetValue<string>()}'.",
                                "General");
                            continue;
                        }

                        list.Add(new ReportDefinition
                        {
                            Server = row.Cell(1).GetValue<string>().Trim(),
                            Database = row.Cell(2).GetValue<string>().Trim(),
                            User = row.Cell(3).GetValue<string>().Trim(),
                            Password = row.Cell(4).GetValue<string>().Trim(),
                            RptPath = row.Cell(5).GetValue<string>().Trim(),
                            ToEmail = row.Cell(6).GetValue<string>().Trim(),
                            CcEmail = row.Cell(7).GetValue<string>().Trim(),
                            ScheduledTime = scheduled,
                            Subject = row.Cell(9).GetValue<string>().Trim(),   
                            Body = row.Cell(10).GetValue<string>().Trim(),
                            PdfName = row.Cell(11).GetValue<string>().Trim()

                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to read schedule Excel.", "General", ex);
            }
            return list;
        }

        // Accepts "HH:mm", "H:mm", or a real Excel time/datetime cell
        private static bool TryParseTime(IXLCell cell, out TimeSpan result)
        {
            result = TimeSpan.Zero;

            // If Excel stored it as a real time/datetime
            try
            {
                if (cell.DataType == XLDataType.DateTime)
                {
                    result = cell.GetDateTime().TimeOfDay;
                    return true;
                }
            }
            catch { /* fall through to text parse */ }

            string text = cell.GetValue<string>().Trim();
            if (TimeSpan.TryParse(text, out result))
                return true;

            DateTime dt;
            if (DateTime.TryParse(text, out dt))
            {
                result = dt.TimeOfDay;
                return true;
            }
            return false;
        }

        // ---- Per-report run (returns true only if PDF made AND email sent) ----
        private bool RunReport(ReportDefinition report)
        {
            Logger.Log($"===== {report.Key} FLASH REPORT STARTED =====", report.Key);
            try
            {
                string pdfPath = ExportCrystalReport(report);
                Logger.Log($"{report.Key}: PDF generated: {pdfPath}", report.Key);

                SendReportEmail(report, pdfPath);
                Logger.Log($"===== {report.Key} FLASH REPORT COMPLETED =====", report.Key);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"===== {report.Key} FLASH REPORT FAILED =====", report.Key, ex);
                return false;
            }
        }

        // ---- Crystal export ----
        private string ExportCrystalReport(ReportDefinition def)
        {
            DateTime reportDate = DateTime.Now.AddDays(-1);

            if (string.IsNullOrEmpty(def.RptPath) || !File.Exists(def.RptPath))
                throw new FileNotFoundException($"Report file not found for {def.Key}: {def.RptPath}");

            string outputDir = ConfigurationManager.AppSettings["ReportOutputPath"];
            if (string.IsNullOrEmpty(outputDir))
                outputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Reports");
            Directory.CreateDirectory(outputDir);

            //string pdfPath = Path.Combine(
            //    outputDir,
            //    $"{def.Key}FlashReport_{reportDate:yyyyMMdd}.pdf");

            string baseName = string.IsNullOrWhiteSpace(def.PdfName)
    ? $"{def.Key}FlashReport"
    : def.PdfName;

            // Strip a trailing .pdf if someone typed it, and clean illegal filename chars
            baseName = baseName.Replace(".pdf", "").Replace(".PDF", "").Trim();
            foreach (char c in Path.GetInvalidFileNameChars())
                baseName = baseName.Replace(c, '_');

            string pdfPath = Path.Combine(
                outputDir,
                $"{baseName}_{reportDate:yyyyMMdd}.pdf");

            if (File.Exists(pdfPath)) File.Delete(pdfPath);

            ReportDocument doc = null;
            try
            {
                doc = new ReportDocument();
                doc.Load(def.RptPath);

                // Log in main report + every subreport/table
                ApplyLogin(doc, def);

                doc.Refresh();

                doc.ExportToDisk(ExportFormatType.PortableDocFormat, pdfPath);
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                    doc.Dispose();
                }
            }

            return pdfPath;
        }

        private void ApplyLogin(ReportDocument doc, ReportDefinition def)
        {
            var ci = new ConnectionInfo
            {
                ServerName = def.Server,   // ODBC DSN name, or host:port for native driver
                DatabaseName = def.Database,
                UserID = def.User,
                Password = def.Password
            };

            // Main report tables
            foreach (Table t in doc.Database.Tables)
            {
                TableLogOnInfo li = t.LogOnInfo;
                li.ConnectionInfo = ci;
                t.ApplyLogOnInfo(li);
            }

            // Subreports (critical for the Foods report)
            foreach (ReportDocument sub in doc.Subreports)
            {
                foreach (Table t in sub.Database.Tables)
                {
                    TableLogOnInfo li = t.LogOnInfo;
                    li.ConnectionInfo = ci;
                    t.ApplyLogOnInfo(li);
                }
            }
        }

        // ---- Email ----
        //        private void SendReportEmail(ReportDefinition report, string pdfPath)
        //        {
        //            DateTime reportDate = DateTime.Now.AddDays(-1);

        //            using (var mail = new MailMessage())
        //            {
        //                mail.From = new MailAddress(ConfigurationManager.AppSettings["FromEmail"]);

        //                foreach (var e in SplitEmails(report.ToEmail)) mail.To.Add(e);
        //                foreach (var e in SplitEmails(report.CcEmail)) mail.CC.Add(e);

        //                mail.Subject = $"{report.Database} - Daily Flash Report - {reportDate:dd/MM/yyyy}";
        //                mail.Body =
        //$@"Dear All,

        //This is a system generated Daily Flash Report of {report.Database}. Please find the attached file. If there is any discrepancy, please contact the SAP Team.

        //Best Regards,
        //SAP Team";

        //                mail.Attachments.Add(new Attachment(pdfPath));
        //                SendViaSmtp(mail);
        //            }
        //        }

        private void SendReportEmail(ReportDefinition report, string pdfPath)
        {
            DateTime reportDate = DateTime.Now.AddDays(-1);
            string dateStr = reportDate.ToString("dd/MM/yyyy");

            // Subject: use Excel value if given, else default
            string subject = string.IsNullOrWhiteSpace(report.Subject)
                ? $"{report.Database} - Daily Flash Report - {dateStr}"
                : report.Subject.Replace("{date}", dateStr).Replace("{db}", report.Database);

            // Body: use Excel value if given, else default
            string body = string.IsNullOrWhiteSpace(report.Body)
                ? $@"Dear All,

This is a system generated Daily Flash Report of {report.Database} for {dateStr}. Please find the attached file. If there is any discrepancy, please contact the SAP Team.

Best Regards,
SAP Team"
                : report.Body.Replace("{date}", dateStr).Replace("{db}", report.Database);

            using (var mail = new MailMessage())
            {
                mail.From = new MailAddress(ConfigurationManager.AppSettings["FromEmail"]);

                foreach (var e in SplitEmails(report.ToEmail)) mail.To.Add(e);
                foreach (var e in SplitEmails(report.CcEmail)) mail.CC.Add(e);

                mail.Subject = subject;
                mail.Body = body;

                mail.Attachments.Add(new Attachment(pdfPath));
                SendViaSmtp(mail);
            }
        }

        private void SendCompletionEmail(List<ReportDefinition> ranReports)
        {
            string to = ConfigurationManager.AppSettings["CompletionReportTo"];
            if (string.IsNullOrWhiteSpace(to))
            {
                Logger.Log("No CompletionReportTo configured; skipping completion email.", "General");
                return;
            }

            DateTime reportDate = DateTime.Now.AddDays(-1);
            string names = string.Join(", ", ranReports.Select(r => r.Database));

            using (var mail = new MailMessage())
            {
                mail.From = new MailAddress(ConfigurationManager.AppSettings["FromEmail"]);
                foreach (var e in SplitEmails(to)) mail.To.Add(e);

                mail.Subject = $"Daily Flash Reports - Completed - {reportDate:dd/MM/yyyy}";
                mail.Body =
$@"Dear All,

The following Daily Flash Report(s) for {reportDate:dd/MM/yyyy} were generated and emailed successfully:
{names}

Best Regards,
SAP Team";

                SendViaSmtp(mail);
            }
        }

        private void SendViaSmtp(MailMessage mail)
        {
            using (var smtp = new SmtpClient(
                ConfigurationManager.AppSettings["SMTPServer"],
                int.Parse(ConfigurationManager.AppSettings["SMTPPort"])))
            {
                smtp.EnableSsl = true;
                smtp.Credentials = new NetworkCredential(
                    ConfigurationManager.AppSettings["SMTPUser"],
                    ConfigurationManager.AppSettings["SMTPPassword"]);
                smtp.Send(mail);
            }
        }

        private static IEnumerable<string> SplitEmails(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) yield break;
            foreach (var e in s.Split(new[] { ';', ',' },
                                      StringSplitOptions.RemoveEmptyEntries))
                yield return e.Trim();
        }

        // ---- Logger ----
        public static class Logger
        {
            private static string GetLogPath()
            {
                string path = ConfigurationManager.AppSettings["LogPath"];
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
                string detail = ex != null
                    ? $"\nException: {ex.Message}\nStacktrace: {ex.StackTrace}"
                    : "";
                Log($"{message} {detail}", dbName, "ERROR");
            }
        }
    }
}