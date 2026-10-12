using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Threading;
using CrystalDecisions.CrystalReports.ViewerObjectModel;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Printing;
using System.Net.Mail;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;
using CrystalDecisions.CrystalReports.Engine;
using System.Runtime.InteropServices;
using CrystalDecisions.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Configuration;
using System.Xml;
using CrystalDecisions.CrystalReports.Engine;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace InvoicePrinting_Sujal
{
    class ClsARCreditNote : Base
    {
         #region VARIABLE DECLARATION

        SAPbouiCOM.Item _ExistItem, _NewItem;
        SAPbouiCOM.Button _Button;
        string _DocEntry = "",_DocNum = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        string _ImportFileName = "";
        private bool flag;
      

        #endregion

         #region CONSTRUCTOR & DISTRUCTOR
        public ClsARCreditNote()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsARCreditNote()
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
                        if (pVal.ItemUID == "1" && (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                        {
                            _DocNum = _Form.DataSources.DBDataSources.Item("ORIN").GetValue("DocNum", 0).ToString();
                            flag = true;
                        }
                        //if (pVal.ItemUID == "BtnPrint")
                        //{
                        //    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();

                        //}

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

                        if (pVal.ItemUID == "1" && pVal.ActionSuccess == true && flag == true)
                        {
                            if (_Form.DataSources.DBDataSources.Item("ORIN").GetValue("U_IsSync", 0).ToString().Trim() == "Y")
                            {

                                IRD(_DocNum);
                                flag = false;
                            }
                        }

                        break;

                    #endregion
                }
            }

            #endregion

        }
        #endregion

        #region Upload data of AR Credit Note on IRD Portal

        public void IRD(string DocNum)
        {
            using (var client = new HttpClient())
            {
                DataSet ds = LoadCompanySetting();
                string WebUserName = ds.Tables[0].Rows[0]["WebUserName"].ToString().Trim();
                string WebPassword = ds.Tables[0].Rows[0]["WebPassword"].ToString().Trim();
                string seller_pan = ds.Tables[0].Rows[0]["seller_pan"].ToString().Trim();
                string buyer_pan = ds.Tables[0].Rows[0]["buyer_pan"].ToString().Trim();
                string fiscal_year = ds.Tables[0].Rows[0]["fiscal_year"].ToString().Trim();
                string InvBaseAddress = ds.Tables[0].Rows[0]["InvBaseAddress"].ToString().Trim();
                string CRNoteresponse = ds.Tables[0].Rows[0]["CRNoteresponse"].ToString().Trim();

                SAPbobsCOM.Recordset _RS = null;
                string oSql = "select Top 1 T0.\"CardCode\",T0.\"CardName\",T0.\"DocTotal\" as \"TotalAmt\",T0.\"VatSum\" as \"TaxAmt\", " +
                              " T0.\"DocTotal\"-T0.\"VatSum\" as \"AmtBeforeTax\",T0.\"DocNum\",case when T1.\"BaseRef\" = '' then T4.\"RefDocNum\" else T1.\"BaseRef\" END as \"BaseRef\",T0.\"U_LD\",T3.\"LicTradNum\",T0.\"DocEntry\" from \"ORIN\" T0  " +
                              "inner join \"RIN1\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                              "inner join OCRD T3 ON T3.\"CardCode\" = T0.\"CardCode\" " +
                               "left outer join RIN21 T4 ON T0.\"DocEntry\" = T4.\"DocEntry\" " +
                              "where T0.\"DocDate\" > '17.07.2022' AND \"DocNum\" = '" + DocNum + "' order by T0.\"DocEntry\" desc ";
                Utilities.ExecuteSQL(ref _RS, oSql);
                if (_RS.RecordCount > 0)
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    BillReturnViewModel p = new BillReturnViewModel
                    {
                        username = WebUserName.ToString(),
                        password = WebPassword.ToString(),//"test@321",
                        seller_pan = seller_pan.ToString(),//"999999999",
                        buyer_pan = _RS.Fields.Item("LicTradNum").Value.ToString().Trim(),//buyer_pan.ToString(),//12345678
                        buyer_name = _RS.Fields.Item("CardName").Value.ToString().Trim(),//"",
                        fiscal_year = fiscal_year.ToString(),//"2074.075",
                        ref_invoice_number = _RS.Fields.Item("BaseRef").Value.ToString().Trim(),//"102",
                        credit_note_date = _RS.Fields.Item("U_LD").Value.ToString().Trim(),//"2074.07.06",
                        credit_note_number = _RS.Fields.Item("DocNum").Value.ToString().Trim(),//"1",
                        reason_for_return = "defect in piece",
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

                    client.BaseAddress = new Uri(InvBaseAddress.ToString());
                    var response = client.PostAsJsonAsync(CRNoteresponse.ToString(), p).Result;
                    if (response.IsSuccessStatusCode)
                    {
                        var responseCode = response.Content.ReadAsStringAsync();
                        if (responseCode.Result == "200")
                        {
                            SAPbobsCOM.Recordset _RS12 = null;
                            string _Qry = "Update \"ORIN\" set \"U_SyncIRD\"='Y',\"U_IsRealTm\" = 'Y',\"U_SyncMsgID\" = '200' Where \"DocNum\"='" + DocNum + "' and \"DocEntry\" = '" + _RS.Fields.Item("DocEntry").Value.ToString().Trim() + "'";
                            Utilities.ExecuteSQL(ref _RS12, _Qry);

                            Console.Write(responseCode.Result);//responseCode 200 means successful
                            Console.ReadLine();
                        }
                        else
                        {
                            SAPbobsCOM.Recordset _RS13 = null;
                            string _Qry1 = "Update \"ORIN\" set \"U_SyncIRD\"='N',\"U_IsRealTm\" = 'N',\"U_SyncMsgID\" = '" + responseCode.Result + "' Where \"DocNum\"='" + DocNum + "' and \"DocEntry\" = '" + _RS.Fields.Item("DocEntry").Value.ToString().Trim() + "'";
                            Utilities.ExecuteSQL(ref _RS13, _Qry1);
                            Console.Write("Error code " + responseCode.Result);
                            Console.ReadLine();
                        }
                    }
                    else
                    { }
                    Console.Write("Error");
                    Console.ReadLine();
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
