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
    class ClsInvoiceSyncIRD : Base
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


        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsInvoiceSyncIRD()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsInvoiceSyncIRD()
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
                            string Fdate = ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value.ToString();
                            string Tdate = ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value.ToString();
                            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific;
                            string bpType = oCombo.Value.Trim();
                            SAPbobsCOM.Recordset _RS = null;

                            Utilities.Message("Please Wait Operation is in process....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            string oSql = "select \"DocEntry\" from OINV where \"DocDate\" > '31.12.2022' AND \"U_IsSync\" = 'Y' and \"U_SyncIRD\" = 'N' and \"DocDate\" between '" + Fdate + "' and '" + Tdate + "'";
                            Utilities.ExecuteSQL(ref _RS, oSql);
                            if (_RS.RecordCount > 0)
                            {
                                for (int i = 0; i < _RS.RecordCount; i++)
                                {
                                    string _DocEntry = _RS.Fields.Item("DocEntry").Value.ToString().Trim();
                                    IRD(_DocEntry);
                                    _RS.MoveNext();
                                }
                            }
                            Utilities.Message("Operation completed sucessfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }

                        break;

                    #endregion
                }
            }

            #endregion

        }
        #endregion

        #region Upload data of invoice on IRD Portal

        public void IRD(string DocEntry)
        {
            DataSet ds = LoadCompanySetting();
            string WebUserName = ds.Tables[0].Rows[0]["WebUserName"].ToString().Trim();
            string WebPassword = ds.Tables[0].Rows[0]["WebPassword"].ToString().Trim();
            string seller_pan = ds.Tables[0].Rows[0]["seller_pan"].ToString().Trim();
            string buyer_pan = ds.Tables[0].Rows[0]["buyer_pan"].ToString().Trim();
            string fiscal_year = ds.Tables[0].Rows[0]["fiscal_year"].ToString().Trim();
            string InvBaseAddress = ds.Tables[0].Rows[0]["InvBaseAddress"].ToString().Trim();
            string Invresponse = ds.Tables[0].Rows[0]["Invresponse"].ToString().Trim();

            using (var client = new HttpClient())
            {

                SAPbobsCOM.Recordset _RS = null;
                string oSql = "select Top 1 OINV.\"CardCode\",OINV.\"CardName\",OINV.\"DocTotal\" as \"TotalAmt\",OINV.\"VatSum\" as \"TaxAmt\",OINV.\"DocTotal\"-OINV.\"VatSum\" as \"AmtBeforeTax\", " +
                              " OINV.\"DocNum\",OINV.\"U_LD\",OCRD.\"LicTradNum\" from OINV inner join OCRD ON OCRD.\"CardCode\" = OINV.\"CardCode\" " +
                              " INNER JOIN NNM1 ON OINV.\"Series\" = NNM1.\"Series\" AND OINV.\"ObjType\" = NNM1.\"ObjectCode\" where OINV.\"DocEntry\" = '" + DocEntry + "' Order by OINV.\"DocEntry\" desc";
                Utilities.ExecuteSQL(ref _RS, oSql);
                string aa = _RS.Fields.Item("U_LD").Value.ToString();
                if (_RS.RecordCount > 0)
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    BillViewModel p = new BillViewModel
                    {
                        username = WebUserName.ToString(),//"Test_CBMS",
                        password = WebPassword.ToString(),//"test@321",
                        seller_pan = seller_pan.ToString(),//"999999999",
                        buyer_pan = _RS.Fields.Item("LicTradNum").Value.ToString().Trim(),//buyer_pan.ToString(),//12345678
                        buyer_name = _RS.Fields.Item("CardName").Value.ToString().Trim(),//"test data"
                        fiscal_year = fiscal_year.ToString(),//"2074.075",
                        invoice_number = _RS.Fields.Item("DocNum").Value.ToString().Trim(),//"SB-017-999",
                        invoice_date = _RS.Fields.Item("U_LD").Value.ToString().Trim(),//"2074.07.06",
                        total_sales = Convert.ToDouble(_RS.Fields.Item("TotalAmt").Value.ToString().Trim()),//1130,
                        taxable_sales_vat = Convert.ToDouble(_RS.Fields.Item("AmtBeforeTax").Value.ToString().Trim()),//1000,
                        vat = Convert.ToDouble(_RS.Fields.Item("TaxAmt").Value.ToString().Trim()),//130,
                        excisable_amount = 0,
                        excise = 0,
                        taxable_sales_hst = 0,
                        hst = 0,
                        amount_for_esf = 0,
                        esf = 0,
                        export_sales = 0,
                        tax_exempted_sales = 0,
                        isrealtime = true,
                        datetimeClient = DateTime.Now
                    };

                    string jsonSent = Newtonsoft.Json.JsonConvert.SerializeObject(p);

                    //client.BaseAddress = new Uri(InvBaseAddress.ToString());
                    //var response = client.PostAsJsonAsync(Invresponse.ToString(), p).Result;

                    //string responseBody = response.Content.ReadAsStringAsync().Result;

                    //client.BaseAddress = new Uri("http://202.166.207.75:9050");
                    //var response = client.PostAsJsonAsync("api/bill", p).Result;

                    client.BaseAddress = new Uri(InvBaseAddress.ToString());
                    var response = client.PostAsJsonAsync(Invresponse.ToString(), p).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        var responseCode = response.Content.ReadAsStringAsync();
                        if (responseCode.Result == "200")
                        {
                            SAPbobsCOM.Recordset _RS12 = null;
                            string _Qry = "Update \"OINV\" set \"U_SyncIRD\"='Y',\"U_IsRealTm\" = 'N',\"U_SyncMsgID\" = '200' Where \"DocEntry\"='" + DocEntry + "'";
                            Utilities.ExecuteSQL(ref _RS12, _Qry);

                            Console.Write(responseCode.Result);//responseCode 200 means successful
                            Console.ReadLine();
                        }
                        else
                        {
                            SAPbobsCOM.Recordset _RS13 = null;
                            string _Qry1 = "Update \"OINV\" set \"U_SyncIRD\"='N',\"U_IsRealTm\" = 'N',\"U_SyncMsgID\" = '" + responseCode.Result + "' Where \"DocEntry\"='" + DocEntry + "'";
                            Utilities.ExecuteSQL(ref _RS13, _Qry1);
                            Console.Write("Error code " + responseCode.Result);
                            Console.ReadLine();
                        }
                    }
                    else
                    {
                        Console.Write("Error");
                        Console.ReadLine();
                    }
                }
            }
        }

        #endregion

        #region LOAD COMPANY SETTINGS
        public DataSet LoadCompanySetting()
        {
            XmlDataDocument xmlDoc = new XmlDataDocument();
            DataSet dataSet = new DataSet();
            dataSet.ReadXml("AppConfig.xml");
            return dataSet;

        }
        #endregion
        
    }
}
