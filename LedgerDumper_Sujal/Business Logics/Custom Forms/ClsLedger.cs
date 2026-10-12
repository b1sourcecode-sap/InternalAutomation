using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.CrystalReports.ViewerObjectModel;
using CrystalDecisions.Shared;
using CrystalDecisions.Shared;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Net.Security;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Outlook = Microsoft.Office.Interop.Outlook;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Colors;
namespace InvoicePrinting_Sujal
{
    class ClsLedger : Base
    {
        #region VARIABLE DECLARATION

        SAPbouiCOM.Item _ExistItem, _NewItem;
        SAPbouiCOM.Button _Button;
        string _DocEntry = "", _DocNum = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        string _ImportFileName = "";
        private bool flag;
        SAPbouiCOM.ComboBox _ComboBox;
        private static readonly HttpClient client = new HttpClient();
        private static string _slCookieHeader = "";


        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsLedger()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsLedger()
        {
        }
        #endregion

        #region ITEM EVENT
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            #region BEFORE ACTION ACTION
            if (pVal.Before_Action == true)
            {
                switch (pVal.EventType)
                {
                    #region ITEM PRESS EVENT
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "3")
                        {
                            string Fdate = ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString().Trim();
                            string Tdate = ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value.ToString().Trim();
                            if (Fdate == "" || Tdate == "")
                            {
                                Utilities.ShowWarningMessage("From date or To Date can not left blank");
                                BubbleEvent = false;
                                return;
                            }
                            // You can hardcode this or use a FolderBrowserDialog

                        }
                        break;
                    #endregion


                }
            }
            #endregion

            #region After ACTION ACTION
            else if (pVal.Before_Action == false)
            {

                switch (pVal.EventType)
                {
                    #region ITEM PRESS EVENT

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        if (pVal.ItemUID == "3" && pVal.ActionSuccess == true)// && flag == true
                        {
                            //string Fdate = ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString();
                            //string Tdate = ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value.ToString();
                            //SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific;
                            //string bpType = oCombo.Value.Trim();
                            //SAPbobsCOM.Recordset _RS = null;
                            //Utilities.Message("Please Wait Operation is in process....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            //string baseFolder = @"C:\SAP_Exports\Ledgers_" + DateTime.Now.ToString("yyyyMMdd");
                            //ExecuteFullDump(Fdate, Tdate, bpType, baseFolder);
                            //Utilities.Message("Operation completed sucessfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            GetReportBinary("RCRI0029");


                        }

                        break;

                    #endregion
                }
            }

            #endregion

        }
        #endregion

        

        public void ExecuteFullDump(string fromDate, string toDate, string bpType, string targetFolder)
        {
            SAPbobsCOM.Recordset rsBP = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                string typeFolderName = (bpType == "C") ? "Customers" : "Vendors";
                string finalPath = Path.Combine(targetFolder, typeFolderName);

                // 2. Create directory if not exists
                if (!Directory.Exists(finalPath)) Directory.CreateDirectory(finalPath);

                // HANA Query to get all Business Partners
                rsBP.DoQuery($@"SELECT ""CardCode"", ""CardName"" FROM OCRD WHERE ""frozenFor"" = 'N' and ""CardType"" = '{bpType}' ORDER BY ""CardCode"" ");

                int count = 0;
                while (!rsBP.EoF)
                {
                    string cardCode = rsBP.Fields.Item("CardCode").Value.ToString();
                    string cardName = rsBP.Fields.Item("CardName").Value.ToString();

                    Utilities.Application.SBO_Application.SetStatusBarMessage($"Processing {count + 1}: {cardCode}", BoMessageTime.bmt_Short, false);

                    GenerateSinglePDF(cardCode, cardName, fromDate, toDate, finalPath, bpType);

                    count++;
                    rsBP.MoveNext();
                }
            }
            finally
            {
                if (rsBP != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsBP);
            }
        }
        private string GetTransTypeTitle(int transType)
        {
            switch (transType)
            {
                case 13: return "AR Invoice";
                case 14: return "AR Credit Memo";
                case 24: return "Incoming Payment";
                case 18: return "AP Invoice";
                case 19: return "AP Credit Memo";
                case 46: return "Outgoing Payment";
                case 30: return "Journal Entry";
                default: return "Trans (" + transType + ")";
            }
        }

        private void GenerateSinglePDF(string cardCode, string cardName, string fDate, string tDate, string folder, string bpType)
        {
            // 1. Create a safe timestamp (e.g., 20231027_143005)
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            // 2. Clean the name as before
            string safeName = string.Join("_", cardName.Split(Path.GetInvalidFileNameChars()));
            // 3. Combine them into the final path
            string filePath = Path.Combine(folder, $"{cardCode}_{safeName}_{timestamp}.pdf");
            string compName = Utilities.Application.Company.CompanyName;
            string bpLabel = (bpType == "C") ? "Customer" : "Supplier";
            // Date Formatting Logic (YYYYMMDD -> DD/MM/YYYY)
            string FormatDate(string d) => d.Length == 8 ? $"{d.Substring(6, 2)}/{d.Substring(4, 2)}/{d.Substring(0, 4)}" : d;
            string dispFrom = FormatDate(fDate);
            string dispTo = FormatDate(tDate);

            using (iText.Kernel.Pdf.PdfWriter writer = new iText.Kernel.Pdf.PdfWriter(filePath))
            {
                using (iText.Kernel.Pdf.PdfDocument pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                {
                    iText.Layout.Document doc = new iText.Layout.Document(pdf);
                    doc.SetMargins(20, 20, 20, 20);

                    iText.Kernel.Font.PdfFont boldFont = iText.Kernel.Font.PdfFontFactory.CreateFont(iText.IO.Font.Constants.StandardFonts.HELVETICA_BOLD);
                    iText.Kernel.Font.PdfFont italicFont = iText.Kernel.Font.PdfFontFactory.CreateFont(iText.IO.Font.Constants.StandardFonts.HELVETICA_OBLIQUE);

                    // --- HEADER ---
                    float[] headerWidths = { 60, 40 };
                    iText.Layout.Element.Table headerTable = new iText.Layout.Element.Table(iText.Layout.Properties.UnitValue.CreatePercentArray(headerWidths)).UseAllAvailableWidth().SetBorder(iText.Layout.Borders.Border.NO_BORDER);

                    iText.Layout.Element.Cell leftCell = new iText.Layout.Element.Cell().SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                    leftCell.Add(new iText.Layout.Element.Paragraph(compName).SetFont(boldFont).SetFontSize(18));
                    leftCell.Add(new iText.Layout.Element.Paragraph("General Ledger").SetFont(boldFont).SetFontSize(10));
                    headerTable.AddCell(leftCell);

                    iText.Layout.Element.Cell rightCell = new iText.Layout.Element.Cell().SetBorder(iText.Layout.Borders.Border.NO_BORDER).SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT);
                    rightCell.Add(new iText.Layout.Element.Paragraph($"Date: {DateTime.Now:dd/MM/yyyy}").SetFontSize(9));
                    rightCell.Add(new iText.Layout.Element.Paragraph($"Time: {DateTime.Now:h:mm tt}").SetFontSize(9));
                    rightCell.Add(new iText.Layout.Element.Paragraph("VAT Number: 301695716").SetFontSize(9));
                    headerTable.AddCell(rightCell);
                    doc.Add(headerTable);

                    // --- PROPERTIES ---
                    doc.Add(new iText.Layout.Element.Paragraph($"Cross-Section Properties:  Posting Date  From: {dispFrom}  To: {dispTo}").SetFontSize(9).SetMarginTop(10));
                    doc.Add(new iText.Layout.Element.Paragraph($"{bpLabel}:  {cardCode}   {cardName}").SetFont(boldFont).SetFontSize(10));

                    double openBal = GetOpeningBalance(cardCode, fDate);
                    doc.Add(new iText.Layout.Element.Paragraph($"Opening Balance: {openBal:N2}").SetFont(italicFont).SetFontSize(9).SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT));

                    // --- TABLE ---
                    float[] columnWidths = { 12, 14, 10, 28, 12, 12, 12 };
                    iText.Layout.Element.Table table = new iText.Layout.Element.Table(iText.Layout.Properties.UnitValue.CreatePercentArray(columnWidths)).UseAllAvailableWidth();

                    string[] headers = { "Date", "Type", "Doc No.", "Remarks", "Debit", "Credit", "Balance" };
                    foreach (var h in headers)
                    {
                        table.AddHeaderCell(new iText.Layout.Element.Cell().Add(new iText.Layout.Element.Paragraph(h)).SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY).SetFont(boldFont).SetFontSize(9).SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER));
                    }

                    SAPbobsCOM.Recordset rsTrans = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    try
                    {
                        string sql = $@"SELECT T0.""RefDate"", T0.""TransType"", T0.""BaseRef"", T0.""Memo"", T1.""Debit"", T1.""Credit"" 
                                FROM JDT1 T1 INNER JOIN OJDT T0 ON T1.""TransId"" = T0.""TransId"" 
                                WHERE T1.""ShortName"" = '{cardCode}' AND T0.""RefDate"" BETWEEN '{fDate}' AND '{tDate}' 
                                ORDER BY T0.""RefDate""";
                        rsTrans.DoQuery(sql);

                        double runningBal = openBal;
                        while (!rsTrans.EoF)
                        {
                            double dr = Convert.ToDouble(rsTrans.Fields.Item("Debit").Value);
                            double cr = Convert.ToDouble(rsTrans.Fields.Item("Credit").Value);
                            runningBal += (dr - cr);

                            table.AddCell(new iText.Layout.Element.Cell().Add(new iText.Layout.Element.Paragraph(Convert.ToDateTime(rsTrans.Fields.Item("RefDate").Value).ToString("dd/MM/yyyy"))).SetFontSize(8));
                            table.AddCell(new iText.Layout.Element.Cell().Add(new iText.Layout.Element.Paragraph(GetTransTypeTitle(Convert.ToInt32(rsTrans.Fields.Item("TransType").Value)))).SetFontSize(8));
                            table.AddCell(new iText.Layout.Element.Cell().Add(new iText.Layout.Element.Paragraph(rsTrans.Fields.Item("BaseRef").Value.ToString())).SetFontSize(8));
                            table.AddCell(new iText.Layout.Element.Cell().Add(new iText.Layout.Element.Paragraph(rsTrans.Fields.Item("Memo").Value.ToString())).SetFontSize(8));
                            table.AddCell(new iText.Layout.Element.Cell().Add(new iText.Layout.Element.Paragraph(dr > 0 ? dr.ToString("N2") : "")).SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT).SetFontSize(8));
                            table.AddCell(new iText.Layout.Element.Cell().Add(new iText.Layout.Element.Paragraph(cr > 0 ? cr.ToString("N2") : "")).SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT).SetFontSize(8));
                            table.AddCell(new iText.Layout.Element.Cell().Add(new iText.Layout.Element.Paragraph(runningBal.ToString("N2"))).SetFont(boldFont).SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT).SetFontSize(8));
                            rsTrans.MoveNext();
                        }
                        doc.Add(table);

                        doc.Add(new iText.Layout.Element.Paragraph($"Closing Balance: {runningBal:N2}").SetFont(boldFont).SetMarginTop(10).SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT));

                        // Footer text removed from here
                    }
                    finally
                    {
                        if (rsTrans != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsTrans);
                        doc.Close();
                    }
                }
            }
        }

        private double GetOpeningBalance(string cardCode, string date)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string sql = $@"SELECT SUM(T1.""Debit"" - T1.""Credit"") as ""Bal"" 
                            FROM JDT1 T1 INNER JOIN OJDT T0 ON T1.""TransId"" = T0.""TransId"" 
                            WHERE T1.""ShortName"" = '{cardCode}' AND T0.""RefDate"" < '{date}'";
            rs.DoQuery(sql);
            double bal = Convert.ToDouble(rs.Fields.Item("Bal").Value);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            return bal;
        }
        #region LOAD COMPANY SETTINGS
        public DataSet LoadCompanySetting()
        {
            XmlDataDocument xmlDoc = new XmlDataDocument();
            DataSet dataSet = new DataSet();
            dataSet.ReadXml("AppConfig.xml");
            return dataSet;

        }
        #endregion

        public async Task GetReportBinary(string layoutCode)
        {
            // 1. Get Session ID
           string cookie = LoginToServiceLayer("https://172.16.1.201:50000/b1s/v2", "UAT_DAIRY", "manager", "sap@9921");

            // 2. Setup PDF Request
            string url = "https://172.16.1.201:50000/b1s/v2/ReportLayoutsService_GetReportLayout";
            string json = "{\"ReportLayoutParams\": {\"LayoutCode\": \"" + layoutCode + "\"}}";

            // 3. Define Path on C Drive
            string savePath = @"C:\SAP_Logs\GeneratedReports\Report_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".pdf";
            if (!Directory.Exists(@"C:\SAP_Logs\GeneratedReports")) Directory.CreateDirectory(@"C:\SAP_Logs\GeneratedReports");

            // 4. Run download using Header authentication
            DownloadBinaryWithHeader(url, json, cookie, savePath);

            Utilities.Message("Report saved successfully using Header Auth", SAPbouiCOM.BoStatusBarMessageType.smt_Success);




            //// 1. CREATE A COOKIE CONTAINER TO HOLD THE SESSION
            //CookieContainer cookies = new CookieContainer();
            //HttpClientHandler handler = new HttpClientHandler
            //{
            //    CookieContainer = cookies,
            //    ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true
            //};

            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            //ServicePointManager.Expect100Continue = false;

            //string folderPath = @"C:\SAP_Logs\GeneratedReports";
            //string fullPath = Path.Combine(folderPath, $"Report_{layoutCode}_{DateTime.Now:yyyyMMdd_HHmm}.pdf");
            //if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);

            //// 2. USE THE HANDLER IN THE CLIENT
            //using (var client = new HttpClient(handler))
            //{
            //    // --- STEP 1: LOGIN (This sets the cookie in the container) ---
            //    var loginRequest = new HttpRequestMessage(HttpMethod.Post, "https://172.16.1.201:50000/b1s/v2/Login");
            //    string loginBody = "{\"CompanyDB\":\"UAT_DAIRY\",\"UserName\":\"manager\",\"Password\":\"sap@9921\"}";
            //    loginRequest.Content = new StringContent(loginBody, Encoding.UTF8, "application/json");
            //    loginRequest.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");

            //    var loginResponse = await client.SendAsync(loginRequest);

            //    if (loginResponse.IsSuccessStatusCode)
            //    {
            //        // --- STEP 2: GET THE PDF (The cookie is now sent automatically) ---
            //        // Changed from GetReportLayout to GetPDF to get the actual binary file
            //        var reportRequest = new HttpRequestMessage(HttpMethod.Post, "https://172.16.1.201:50000/b1s/v2/ReportLayoutsService_GetReportLayout");

            //        string reportParams = $"{{\"ReportLayoutPrintParams\": {{\"LayoutCode\": \"{layoutCode}\"}}}}";
            //        reportRequest.Content = new StringContent(reportParams, Encoding.UTF8, "application/json");
            //        reportRequest.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");

            //        var reportResponse = await client.SendAsync(reportRequest);

            //        if (reportResponse.IsSuccessStatusCode)
            //        {
            //            // --- STEP 3: SAVE THE PDF BYTES ---
            //            using (var stream = await reportResponse.Content.ReadAsStreamAsync())
            //            using (var fileStream = File.Create(fullPath))
            //            {
            //                await stream.CopyToAsync(fileStream);
            //            }

            //            Utilities.Message("PDF Saved Successfully!", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            //            System.Diagnostics.Process.Start(fullPath);
            //        }
            //        else
            //        {
            //            // If it fails here, read the error message
            //            string error = await reportResponse.Content.ReadAsStringAsync();
            //            MessageBox.Show("Report Error: " + error);
            //        }
            //    }
            //    else
            //    {
            //        Utilities.Message("Login Failed", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //    }
            //}
        }


        private string LoginToServiceLayer(string slUrl, string targetDB, string slUser, string slPass)
        {
            try
            {
                System.Net.ServicePointManager.Expect100Continue = false;
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                System.Net.ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;

                // Use v2 consistently if you are calling v2 reports
                string loginUrl = slUrl.TrimEnd('/') + "/Login";

                var request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(loginUrl);
                request.Method = "POST";
                request.ContentType = "application/json";

                var loginParams = new { CompanyDB = targetDB, UserName = slUser, Password = slPass };
                string rawJson = new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(loginParams);
                byte[] byteArray = System.Text.Encoding.UTF8.GetBytes(rawJson);
                request.ContentLength = byteArray.Length;

                using (var dataStream = request.GetRequestStream())
                    dataStream.Write(byteArray, 0, byteArray.Length);

                using (var response = (System.Net.HttpWebResponse)request.GetResponse())
                {
                    using (var reader = new System.IO.StreamReader(response.GetResponseStream()))
                    {
                        string jsonBody = reader.ReadToEnd();
                        // Extract ONLY the SessionId GUID string
                        var dict = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<Dictionary<string, object>>(jsonBody);
                        return dict["SessionId"].ToString();
                    }
                }
            }
            catch (Exception ex) { throw new Exception("Login Failed: " + ex.Message); }
        }

        private void DownloadBinaryWithHeader(string url, string json, string sessionId, string savePath)
        {
            try
            {
                var request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
                request.Method = "POST";
                request.ContentType = "application/json; charset=utf-8";
                request.Accept = "application/json"; 
                request.Headers.Add("Prefer", "return=representation");
                if (!string.IsNullOrEmpty(sessionId))
                {
                    // CRITICAL: Format as 'B1SESSION=ID' only. No paths, no HttpOnly.
                    request.Headers.Add("Cookie", "B1SESSION=" + sessionId);
                }

                byte[] byteArray = System.Text.Encoding.UTF8.GetBytes(json);
                request.ContentLength = byteArray.Length;

                using (System.IO.Stream dataStream = request.GetRequestStream())
                    dataStream.Write(byteArray, 0, byteArray.Length);

                using (var response = (System.Net.HttpWebResponse)request.GetResponse())
                {
                    using (var responseStream = response.GetResponseStream())
                    using (var fileStream = System.IO.File.Create(savePath))
                    {
                        responseStream.CopyTo(fileStream);
                    }
                }
            }

            catch (System.Net.WebException ex)
            {
                if (ex.Response != null)
                {
                    using (var reader = new System.IO.StreamReader(ex.Response.GetResponseStream()))
                    {
                        string errorBody = reader.ReadToEnd();
                        // This will show the actual SAP error code and message
                        MessageBox.Show("SAP Detailed Error: " + errorBody);
                    }
                }
                throw;
            }
        }

    }
}
