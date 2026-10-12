using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.IO;

namespace InvoicePrinting_Sujal
{
    class ClsApCreditMemo : Base
    {

        #region VARIABLE DECLARATION
        private SAPbobsCOM.Recordset _RS;
        private string oSql, mobile, SqlQry, DocNum, DocNum1, CustRefNo, DocuTotal, sDocDate, ChkString;
        private bool flag;
        public SAPbouiCOM.EditText _EditText;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsApCreditMemo()
            : base()
        {
            _Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsApCreditMemo()
        {
        }
        #endregion

        #region ITEM EVENT
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            #region BEFORE ACTION ACTION TRUE
            if (pVal.Before_Action == true)
            {
                switch (pVal.EventType)
                {
                    #region ITEM PRESS EVENT
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            DocNum = "";
                            _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                            _EditText = (SAPbouiCOM.EditText)_Form.Items.Item("8").Specific;
                            sDocDate = Utilities.ToDate(((SAPbouiCOM.EditText)this._Form.Items.Item("10").Specific).Value).ToString("dd-MM-yyyy");
                            DocNum = _EditText.Value.ToString();
                            flag = true;

                        }

                        break;
                    #endregion
                }
            }
            else if (pVal.Before_Action == false)
            {

                switch (pVal.EventType)
                {
                    #region ITEM PRESS EVENT
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && pVal.ActionSuccess == true && flag == true)
                        {
                            oSql = "select O.DocNum,O.CardCode,O.CardName,O.DocDate,O.DocTotal,RD.Cellular,O.NumAtCard from ORPC O " +
                                 "inner join OCRD RD ON O.CardCode = RD.CardCode where DocType = 'I' and O.DocNum = '" + DocNum + "'";

                            Utilities.ExecuteSQL(ref _RS, oSql);
                            DocNum1 = _RS.Fields.Item(0).Value.ToString();
                            DocuTotal = _RS.Fields.Item(4).Value.ToString();
                            mobile = _RS.Fields.Item(5).Value.ToString();
                            CustRefNo = _RS.Fields.Item(6).Value.ToString();

                            string demouser = "20140275";
                            string DM = "infoRR";
                            string Pwd = "manish";
                            string Msg = "This is to inform you that we have debited your account by " + DocuTotal + " against your invoice no " + CustRefNo + " ";
                           
                            SendSMS(demouser, Pwd, mobile, Msg, DM);

                            flag = false;

                        }
                        break;

                    #endregion
                }
            }

            #endregion

        }
        #endregion

        #region SendSMS Method

        public void SendSMS(string demouser, string Pwd, string mobile, string Msg, string DM)
        {
            if (mobile != string.Empty)
            {
                try
                {

                    System.Net.ServicePointManager.Expect100Continue = false;
                    HttpWebRequest myReq = (HttpWebRequest)WebRequest.Create("http://103.247.98.91/API/SendMsg.aspx?uname=" + demouser + "&pass=" + Pwd + "&send=" + DM + "&dest=" + mobile + "&msg=" + Msg + "");

                    myReq.Headers.Clear();
                    myReq.Method = "POST";
                    myReq.KeepAlive = false;
                    myReq.ProtocolVersion = HttpVersion.Version11;
                    myReq.ContentType = "text/xml";
                    myReq.Proxy = null;
                    myReq.Credentials = null;
                    using (StreamWriter sendingData = new StreamWriter(myReq.GetRequestStream()))
                    {

                        sendingData.Write(Msg);
                        sendingData.Flush();
                        sendingData.Close();
                    }

                    HttpWebResponse myResponse = (HttpWebResponse)myReq.GetResponse();
                    StreamReader responseStream = new StreamReader(myResponse.GetResponseStream());
                    Msg = responseStream.ReadToEnd();

                    responseStream.Close();
                    myResponse.Close();


                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }


            }

        }

        #endregion

    }
}
