using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using SAPbobsCOM;
//using SAPbouiCOM;
using System.Data;
using System.Data.SqlClient;
using System.Data.Sql;
using Microsoft.ApplicationBlocks.Data;
using System.Xml;
using Sap.Data.Hana;
using System.IO;

namespace InvoicePrinting_Sujal
{
    class ClsGoodsReturn : Base
    {

            #region VARIABLE DECLARATION

            private static HanaConnection _conn;
            string _LastErrorMsg = " ", strcon = "";
            private static HanaDataAdapter _da;
            private static HanaParameter _parm1;
            private static HanaParameter _parm2;
            SAPbouiCOM.Item _ExistItem, _NewItem;
            SAPbouiCOM.Button _Button;
            string _DocEntry = "", _DocNum = "";
            CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
            string ReportPath = string.Empty;
            string _ImportFileName = "";
            private bool flag;
            private Hashtable _lines = new Hashtable();
            SAPbobsCOM.Recordset _RS2 = null;
            public string oSql2 = string.Empty;
            public string SalesReturnNo = string.Empty, BranchCd = string.Empty;
            SAPbobsCOM.Recordset _Rs1 = null, _RSGRPO = null;

            #endregion

            #region CONSTRUCTOR & DISTRUCTOR

            public ClsGoodsReturn()
                : base()
            {
                //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
                //
                // TODO: Add constructor logic here
                //
            }
            ~ClsGoodsReturn()
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
                            if (pVal.ItemUID == "SR")
                            {
                                _DocEntry = _Form.DataSources.DBDataSources.Item("ORPD").GetValue("DocEntry", 0).ToString();
                                SAPbobsCOM.Recordset _RS = null, _RS12 = null;
                                string PostingDate = "", DeliveryDate = "", DocumentDate = "",LocalDate = "";
                                _Form.Items.Item("SR").Enabled = false;
                                string oSql = "select ORPD.\"CardCode\",ORPD.\"DocDueDate\",ORPD.\"DocDate\",ORPD.\"TaxDate\",ORPD.\"DocEntry\",ORPD.\"DocNum\",ORPD.\"U_RecevingBranch\",ORPD.\"U_CRNote\",OCRD.\"U_GRPOBPCode\",ORPD.\"U_LD\" from ORPD inner join OCRD ON ORPD.\"CardCode\" = OCRD.\"CardCode\" where ORPD.\"DocEntry\" = '" + _DocEntry + "'  order by \"DocDate\"";
                                Utilities.ExecuteSQL(ref _RS, oSql);
                                _DocEntry = _RS.Fields.Item("DocEntry").Value.ToString();
                                _DocNum = _RS.Fields.Item("DocNum").Value.ToString();
                                PostingDate = _RS.Fields.Item("DocDate").Value.ToString();
                                LocalDate = _RS.Fields.Item("U_LD").Value.ToString();
                                DeliveryDate = _RS.Fields.Item("DocDueDate").Value.ToString();
                                DocumentDate = _RS.Fields.Item("TaxDate").Value.ToString();
                                BranchCd = _RS.Fields.Item("U_RecevingBranch").Value.ToString();
                                SalesReturnNo = _Form.DataSources.DBDataSources.Item("ORPD").GetValue("U_CRNote", 0).ToString();
                                string BPCode = _RS.Fields.Item("U_GRPOBPCode").Value.ToString();
                                if (_RS.Fields.Item("U_GRPOBPCode").Value.ToString() != "" && SalesReturnNo == "")
                                {

                                    CreateSR(Convert.ToInt32(_DocNum), Convert.ToInt32(_DocEntry), BPCode, BranchCd, PostingDate, DeliveryDate, DocumentDate, LocalDate);
                                    flag = false;
                                }
                                else
                                {
                                    _Form.Items.Item("SR").Enabled = true;
                                    Utilities.Message("This customer/dealer is not eligible for intercompany transaction", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                    return;
                                }
                                flag = true;
                            }

                            break;
                        case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                            if (pVal.ItemUID == "U_RecevingWHS" && pVal.CharPressed == 9)
                            {
                                if (((SAPbouiCOM.EditText)_Form.Items.Item("U_RecevingWHS").Specific).Value == "")
                                {
                                    SAPbobsCOM.Recordset _RS = null, _RS1 = null;
                                    string _DBName = string.Empty;
                                    string StrQuery = string.Empty;
                                    string StrSchema = string.Empty;

                                    StrQuery = "select T0.\"WhsCode\",T0.\"WhsName\",T0.\"BPLid\",T1.\"BPLName\" from OWHS T0 left outer join OBPL T1 ON T0.\"BPLid\" = T1.\"BPLId\"";
                                    Utilities.ExecuteSQL(ref _RS1, StrQuery);

                                    this._Object = new ClsChooseFromList();
                                    ((ClsChooseFromList)this._Object).FormTitle = "Warehouse";
                                    ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref BubbleEvent, StrQuery, strcon, false, _DBName);

                                }
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
                            if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                SalesReturnNo = _Form.DataSources.DBDataSources.Item("ORPD").GetValue("U_CRNote", 0).ToString();
                                if (SalesReturnNo != "")
                                    _Form.Items.Item("SR").Enabled = false;
                                else
                                    _Form.Items.Item("SR").Enabled = true;
                            }

                            break;

                        #endregion

                        #region et_VALIDATE

                        case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                            if (pVal.ItemUID == "4")
                            {
                                SAPbobsCOM.Recordset _RS1 = null;
                                if (((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value != "")
                                {
                                    SalesReturnNo = _Form.DataSources.DBDataSources.Item("ORPD").GetValue("U_CRNote", 0).ToString();
                                    if (SalesReturnNo != "")
                                        _Form.Items.Item("SR").Enabled = false;
                                    else
                                        _Form.Items.Item("SR").Enabled = true;
                                }
                            }
                            break;

                        #endregion

                    }
                }

                #endregion

            }

            #endregion

            #region Menu_Event

            public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
            {
                try
                {

                    if (pVal.BeforeAction)
                    {

                        switch (pVal.MenuUID)
                        {
                            case Constants.System_Menus.mnu_ADD:
                                SalesReturnNo = _Form.DataSources.DBDataSources.Item("ORPD").GetValue("U_CRNote", 0).ToString();
                                if (SalesReturnNo != "")
                                    _Form.Items.Item("SR").Enabled = false;
                                else
                                    _Form.Items.Item("SR").Enabled = true;

                                break;

                        }
                    }
                    else
                    {
                        switch (pVal.MenuUID)
                        {

                            case Constants.System_Menus.mnu_ADD:
                                SalesReturnNo = _Form.DataSources.DBDataSources.Item("ORPD").GetValue("U_CRNote", 0).ToString();
                                if (SalesReturnNo != "")
                                    _Form.Items.Item("SR").Enabled = false;
                                else
                                    _Form.Items.Item("SR").Enabled = true;

                                break;

                            case Constants.System_Menus.mnu_FIND:
                                SalesReturnNo = _Form.DataSources.DBDataSources.Item("ORPD").GetValue("U_CRNote", 0).ToString();
                                if (SalesReturnNo != "")
                                    _Form.Items.Item("SR").Enabled = false;
                                else
                                    _Form.Items.Item("SR").Enabled = true;

                                break;
                            case Constants.System_Menus.mnu_FIRST:
                                SalesReturnNo = _Form.DataSources.DBDataSources.Item("ORPD").GetValue("U_CRNote", 0).ToString();
                                if (SalesReturnNo != "")
                                    _Form.Items.Item("SR").Enabled = false;
                                else
                                    _Form.Items.Item("SR").Enabled = true;

                                break;
                            case Constants.System_Menus.mnu_LAST:
                                SalesReturnNo = _Form.DataSources.DBDataSources.Item("ORPD").GetValue("U_CRNote", 0).ToString();
                                if (SalesReturnNo != "")
                                    _Form.Items.Item("SR").Enabled = false;
                                else
                                    _Form.Items.Item("SR").Enabled = true;

                                break;
                            case Constants.System_Menus.mnu_NEXT:
                                SalesReturnNo = _Form.DataSources.DBDataSources.Item("ORPD").GetValue("U_CRNote", 0).ToString();
                                if (SalesReturnNo != "")
                                    _Form.Items.Item("SR").Enabled = false;
                                else
                                    _Form.Items.Item("SR").Enabled = true;

                                break;
                            case Constants.System_Menus.mnu_PREVIOUS:
                                SalesReturnNo = _Form.DataSources.DBDataSources.Item("ORPD").GetValue("U_CRNote", 0).ToString();
                                if (SalesReturnNo != "")
                                    _Form.Items.Item("SR").Enabled = false;
                                else
                                    _Form.Items.Item("SR").Enabled = true;

                                break;
                            case Constants.System_Menus.mnu_REFRESH:
                                SalesReturnNo = _Form.DataSources.DBDataSources.Item("ORPD").GetValue("U_CRNote", 0).ToString();
                                if (SalesReturnNo != "")
                                    _Form.Items.Item("SR").Enabled = false;
                                else
                                    _Form.Items.Item("SR").Enabled = true;

                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
                finally
                {
                    Form.Freeze(false);
                }
            }

            #endregion

            #region FUNCTION FOR WRITE IMPORTLOG.LOG FILE

            public static void WriteLogFile(string Message)
            {
                try
                {

                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\Log"))
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\Log");

                    string date = DateTime.Now.Date.Day.ToString() + DateTime.Now.Date.Month.ToString() + DateTime.Now.Date.Year.ToString();
                    FileStream fs = File.Open(System.Windows.Forms.Application.StartupPath + "\\Log\\b1log" + date + ".txt", FileMode.Append, FileAccess.Write);
                    StreamWriter sr = new StreamWriter(fs);
                    sr.WriteLine(Message);
                    sr.Flush();
                    sr.Close();

                }
                catch (Exception ex)
                {

                    throw ex;
                }
            }

            public static void WriteLogFile(string LogFileNameC, string ProcessName, string Message)
            {
                try
                {
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\Log"))
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\Log");

                    string date = DateTime.Now.Date.Day.ToString() + DateTime.Now.Date.Month.ToString() + DateTime.Now.Date.Year.ToString();
                    FileStream fs = File.Open(System.Windows.Forms.Application.StartupPath + "\\Log\\" + LogFileNameC + date + ".txt", FileMode.Append, FileAccess.Write);

                    StreamWriter sr = new StreamWriter(fs);
                    sr.WriteLine(ProcessName + ":    " + Message + DateTime.Now);
                    sr.Flush();
                    sr.Close();
                }
                catch (Exception ex)
                {
                    throw ex;
                }

            }

            #endregion

            #region CreateGRPO

            private void CreateSR(int _DocNum, int _DocEntry, string BPCode, string BranchCode, string PostingDate, string DeliveryDate, string DocumentDate,string _LocalDate)
            {
                try
                {
                    string lineInTrn = "";
                    string MainTbl = "";
                    string LineTbl = "";
                    int skipLine = 0;


                    SAPbobsCOM.BoObjectTypes objType = SAPbobsCOM.BoObjectTypes.oPurchaseReturns;
                    SAPbobsCOM.Documents document;

                    MainTbl = "ORPD";
                    LineTbl = "RPD1";
                    objType = SAPbobsCOM.BoObjectTypes.oPurchaseReturns;
                    Utilities.ShowWarningMessage("Please wait...");

                    string sql = "  SELECT T.\"DiscPrcnt\",T.\"U_RecevingWHS\",T.\"U_RecevingBranch\",T1.\"WhsCode\",T1.\"Price\", T1.\"ItemCode\",T1.\"Dscription\",T1.\"Quantity\",T1.\"LineNum\",T2.\"ManSerNum\",T2.\"ManBtchNum\",T1.\"DocEntry\",T1.\"Project\",T.\"U_InvNo\" "
                                 + String.Format(" FROM {0} T1 INNER JOIN {1} T ON T.\"DocEntry\" =T1.\"DocEntry\" ", LineTbl, MainTbl)
                                 + " Inner join OITM T2 on T1.\"ItemCode\" =T2.\"ItemCode\" "
                                 + String.Format(" where T.\"DocEntry\" ={0} ", _DocEntry)
                                 + " group by T.\"DiscPrcnt\",T.\"U_RecevingWHS\",T.\"U_RecevingBranch\",T1.\"WhsCode\",T1.\"Price\", T1.\"ItemCode\",T1.\"Dscription\",T1.\"Quantity\",T1.\"LineNum\",T2.\"ManSerNum\",T2.\"ManBtchNum\",T1.\"DocEntry\",T1.\"Project\",T.\"U_InvNo\" "
                                 + " order by \"LineNum\" asc ";

                    Utilities.ExecuteSQL(ref _RSGRPO, sql);

                    SAPbobsCOM.Documents _Returns = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                    document = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(objType);
                    document.GetByKey(_DocEntry);

                    _Returns.CardCode = BPCode.ToString();
                    //_GRPO.DocDate = Utilities.ToDate(DateTime.Now.ToString("yyyyMMdd"));

                    DateTime parsedPostingDate = DateTime.Parse(PostingDate);
                    parsedPostingDate.ToString("yyyyMMdd");

                    DateTime parsedPDeliveryDate = DateTime.Parse(DeliveryDate);
                    parsedPDeliveryDate.ToString("yyyyMMdd");

                    DateTime parsedPDocumentDate = DateTime.Parse(DeliveryDate);
                    parsedPDocumentDate.ToString("yyyyMMdd");

                    DateTime LocalDate = DateTime.Parse(_LocalDate);
                    LocalDate.ToString("yyyyMMdd");


                    _Returns.DocDate = Utilities.ToDate(parsedPostingDate.ToString("yyyyMMdd"));
                    _Returns.DocDueDate = Utilities.ToDate(parsedPDeliveryDate.ToString("yyyyMMdd"));
                    _Returns.TaxDate = Utilities.ToDate(parsedPDocumentDate.ToString("yyyyMMdd"));
                    _Returns.UserFields.Fields.Item("U_LD").Value = _LocalDate.ToString();

                    if (BranchCode.ToString().Trim() != "0")
                        _Returns.BPL_IDAssignedToInvoice = Convert.ToInt16(BranchCode.ToString());


                    _Returns.UserFields.Fields.Item("U_GRNum").Value = _DocEntry.ToString();

                    if (_RSGRPO.RecordCount > 0)
                    {
                        lineInTrn = "";
                        
                        _Returns.DiscountPercent = double.Parse(_RSGRPO.Fields.Item("DiscPrcnt").Value.ToString());
                        if (_RSGRPO.RecordCount > 0)
                        {

                            for (int k = 0; k <= _RSGRPO.RecordCount - 1; k++)
                            {
                                _Returns.Lines.SetCurrentLine(k);

                                _Returns.Lines.ItemCode = _RSGRPO.Fields.Item("itemcode").Value.ToString();
                                _Returns.Lines.WarehouseCode = _RSGRPO.Fields.Item("U_RecevingWHS").Value.ToString();
                                _Returns.Lines.Quantity = double.Parse(_RSGRPO.Fields.Item("Quantity").Value.ToString());
                                _Returns.Lines.UnitPrice = double.Parse(_RSGRPO.Fields.Item("Price").Value.ToString());
                                _Returns.Lines.ProjectCode = _RSGRPO.Fields.Item("Project").Value.ToString();
                                _Returns.Lines.BaseType = 13;
                                _Returns.Lines.BaseEntry = Convert.ToInt32(_RSGRPO.Fields.Item("U_InvNo").Value.ToString());
                                _Returns.Lines.BaseLine = Convert.ToInt32(_RSGRPO.Fields.Item("LineNum").Value.ToString());
                                //_Returns.Lines.UserFields.Fields.Item("U_GRNum").Value = _RSGRPO.Fields.Item("DocEntry").Value.ToString();
                               
                                _Returns.Lines.Add();
                                _RSGRPO.MoveNext();
                            }

                            int RetVal = _Returns.Add();
                            if (RetVal == 0)
                            {
                                string objCode = "";
                                Utilities.Application.Company.GetNewObjectCode(out objCode);
                                SAPbobsCOM.Recordset _Rs = null;
                                _Rs1 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string StrQry1 = "select \"DocNum\" from ORIN where \"DocEntry\"  = '" + objCode + "'";
                                Utilities.ExecuteSQL(ref _Rs1, StrQry1);

                                string StrQry = "update ORPD set \"U_CRNote\" = '" + objCode + "' where \"DocEntry\" = '" + _DocEntry + "'";
                                Utilities.ExecuteSQL(ref _Rs, StrQry);
                                WriteLogFile("Log", "Credit Note Created Sucessfully", "  :  SR Entry :" + objCode);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(Utilities.Application.Company);//For releasing the company object
                            }
                            else
                            {
                                _Form.Items.Item("SR").Enabled = true;
                                string GetLastErrorDescription = Utilities.Application.Company.GetLastErrorDescription();
                                WriteLogFile("Log", "Sales Return Creation", "  :  " + GetLastErrorDescription);
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                            }
                        }
                    }
                    Utilities.ShowWarningMessage("Operation completed successfully.");
                }
                catch (Exception ex)
                {
                    Utilities.LogEntry("Error: "
               + ex.Message, "Error in SR : DocEntry "
               + _DocNum + " @ " + System.DateTime.Now.ToString());
                }
            }

            #endregion

            #region LOAD COMPANY SETTINGS

            public static DataSet LoadCompanySetting()
            {
                XmlDataDocument xmlDoc = new XmlDataDocument();
                DataSet dataSet = new DataSet();
                dataSet.ReadXml("AppConfig.xml");
                return dataSet;

            }

            #endregion

            #region CHOOSE FROM LIST EVENT

            public override void ChooseFromList_Event()
            {
                try
                {
                    _Form.Freeze(true);
                    if (this._CFLEvent.ItemUID == "U_RecevingWHS")
                    {
                        ((SAPbouiCOM.EditText)_Form.Items.Item("U_RecevingWHS").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
                        ((SAPbouiCOM.EditText)_Form.Items.Item("U_RecevingBranch").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][2].ToString();
                        ((SAPbouiCOM.EditText)_Form.Items.Item("U_RecevingBranchName").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][3].ToString();

                    }

                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.ToString());
                }
                finally
                {
                    _Form.Freeze(false);
                }
            }

            #endregion

            #region Form Default

            internal void FormDefault()
            {
                try
                {
                    _ExistItem = _Form.Items.Item("2");
                    _NewItem = _Form.Items.Add("SR", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 150;
                    _NewItem.Width = _ExistItem.Width + 20;
                    _NewItem.Top = _ExistItem.Top;
                    _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                    _Button.Caption = "Sales Reruen";

                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
            }

            #endregion

        }

    }

