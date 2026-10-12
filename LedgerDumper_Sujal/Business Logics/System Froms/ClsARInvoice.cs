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
    class ClsARInvoice :Base
    {

        #region VARIABLE DECLARATION

        SAPbouiCOM.Item _ExistItem, _NewItem;
        SAPbouiCOM.Button _Button;
        string _DocEntry = "",_DocNum = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        string _ImportFileName = "";
        private bool flag;
        public string BranchCd = string.Empty,_Series = string.Empty,_WhsCode = string.Empty ;
        SAPbobsCOM.Recordset _Rs1 = null, _RSGRPO = null;
        public string GRPOINVNo = string.Empty;
        string _LastErrorMsg = " ", strcon = "";
        public string GRPONumber = string.Empty;
        SAPbobsCOM.Recordset _RS = null, _RS1 = null;
        public string _DBName = string.Empty;
        public string StrQuery = string.Empty;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsARInvoice()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsARInvoice()
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
                            _DocNum = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocNum", 0).ToString();
                            flag = true;
                        }
                        if (pVal.ItemUID == "BtnPrint")
                        {
                            _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                            
                        }

                        break;
                    #endregion

                    #region et_KEY_DOWN

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "U_RecevingWHS" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("U_RecevingWHS").Specific).Value == "")
                            {
                                
                               

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

                        if (pVal.ItemUID == "GRPO")
                        {
                            #region GRPO Creation

                            _DocNum = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocNum", 0).ToString();
                            //BranchCd = _Form.DataSources.DBDataSources.Item("OINV").GetValue("BPLId", 0).ToString();
                            SAPbobsCOM.Recordset _RS = null, _RS12 = null;
                            string PostingDate = "", DeliveryDate = "", DocumentDate = "";
                            //Utilities.Message("Please wait GRPO creation is in progress....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            _Form.Items.Item("GRPO").Enabled = false;
                            string oSql = "select T0.\"CardCode\",T0.\"DocDueDate\",T0.\"DocDate\",T0.\"TaxDate\",T0.\"DocEntry\",T0.\"DocNum\",T2.\"Remark\",T1.\"U_GRPOBPCode\",T0.\"U_RecevingWHS\",T0.\"U_RecevingBranch\",T0.\"U_RecevingBranchName\" from OINV T0  " +
                                          "inner join OCRD T1 ON T0.\"CardCode\" = T1.\"CardCode\" " +
                                          "inner join NNM1 T2 ON T0.\"Series\" = T2.\"Series\" AND T2.\"ObjectCode\" = 13 " +
                                          "where T0.\"DocEntry\" = '" + _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString() + "'  order by \"DocDate\"";
                            Utilities.ExecuteSQL(ref _RS, oSql);
                            _DocEntry = _RS.Fields.Item("DocEntry").Value.ToString();
                            _DocNum = _RS.Fields.Item("DocNum").Value.ToString();
                            PostingDate = _RS.Fields.Item("DocDate").Value.ToString();
                            DeliveryDate = _RS.Fields.Item("DocDueDate").Value.ToString();
                            DocumentDate = _RS.Fields.Item("TaxDate").Value.ToString();
                            _Series = _RS.Fields.Item("Remark").Value.ToString();
                            string BPCode = _RS.Fields.Item("U_GRPOBPCode").Value.ToString();
                            BranchCd = _RS.Fields.Item("U_RecevingBranch").Value.ToString();
                            _WhsCode = _RS.Fields.Item("U_RecevingWHS").Value.ToString();
                            GRPONumber = _Form.DataSources.DBDataSources.Item("OINV").GetValue("U_GRPONum", 0).ToString();
                            if (_RS.Fields.Item("U_GRPOBPCode").Value.ToString() != "" && GRPONumber == "")
                            {
                                CreateGRPO(Convert.ToInt32(_DocNum), Convert.ToInt32(_DocEntry), BPCode, BranchCd, PostingDate, DeliveryDate, DocumentDate, _Series, _WhsCode);
                            }
                            #endregion
                        }
                        if (pVal.ItemUID == "1" && pVal.ActionSuccess == true && flag == true)
                        {
                            if (_Form.DataSources.DBDataSources.Item("OINV").GetValue("U_IsSync", 0).ToString().Trim() == "Y")
                            {
                                #region SendEmail
                                IRD(_DocNum);
                                // SendEmail(_DocNum);
                                flag = false;
                                _DocNum = "";
                            }
                            #endregion
                        }

                        if (pVal.ItemUID == "BtnPrint")
                        {
                            #region Invoice Printing 
                            try
                            {
                                int PrintNo;
                                SAPbobsCOM.Recordset Rs = null;
                                if (_Form.DataSources.DBDataSources.Item("OINV").GetValue("DocType", 0).ToString().Trim() == "I")
                                {
                                    if (_Form.DataSources.DBDataSources.Item("OINV").GetValue("U_InvoiceType", 0).ToString().Trim() == "Regular Invoice")
                                    {
                                        if (_Form.DataSources.DBDataSources.Item("OINV").GetValue("U_PrintNo", 0).ToString().Trim() == "")
                                        {
                                            PrintNo = 0;

                                            ExportReportToPdfFile();
                                            ExportReportToPdfFile1();
                                            ExportReportToPdfFile2();

                                            string Qry = "Update \"OINV\" set \"U_PrintNo\"='1',\"U_PrintStatus\"='Printed',\"U_PrintDate\"=(select CURRENT_TIMESTAMP from DUMMY),\"U_PrintDTM\"=(select CURRENT_TIME from DUMMY) Where \"DocEntry\"='" + _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString() + "'";
                                            Utilities.ExecuteSQL(ref Rs, Qry);
                                        }
                                        else
                                        {
                                            ExportReportToPdfFile3();
                                            string Qry = "Update \"OINV\" set \"U_PrintNo\"= \"U_PrintNo\" + 1 Where \"DocEntry\"='" + _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString() + "'";
                                            Utilities.ExecuteSQL(ref Rs, Qry);
                                        }

                                    }
                                    else if (_Form.DataSources.DBDataSources.Item("OINV").GetValue("U_InvoiceType", 0).ToString().Trim() == "Fixed Asset Sale")
                                    {
                                        if (_Form.DataSources.DBDataSources.Item("OINV").GetValue("U_PrintNo", 0).ToString().Trim() == "")
                                        {
                                            PrintNo = 0;

                                            FAExportReportToPdfFile();
                                            FAExportReportToPdfFile1();
                                            FAExportReportToPdfFile2();

                                            string Qry = "Update \"OINV\" set \"U_PrintNo\"='1',\"U_PrintStatus\"='Printed',\"U_PrintDate\"=(select CURRENT_TIMESTAMP from DUMMY),\"U_PrintDTM\"=(select CURRENT_TIME from DUMMY) Where \"DocEntry\"='" + _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString() + "'";
                                            Utilities.ExecuteSQL(ref Rs, Qry);
                                        }
                                        else
                                        {
                                            FAExportReportToPdfFile3();
                                            string Qry = "Update \"OINV\" set \"U_PrintNo\"= \"U_PrintNo\" + 1 Where \"DocEntry\"='" + _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString() + "'";
                                            Utilities.ExecuteSQL(ref Rs, Qry);
                                        }
                                    }
                                }
                                else if (_Form.DataSources.DBDataSources.Item("OINV").GetValue("DocType", 0).ToString().Trim() == "S")
                                {
                                    if (_Form.DataSources.DBDataSources.Item("OINV").GetValue("U_PrintNo", 0).ToString().Trim() == "")
                                    {
                                        PrintNo = 0;

                                        ServiceExportReportToPdfFile();
                                        ServiceExportReportToPdfFile1();
                                        ServiceExportReportToPdfFile2();

                                        string Qry = "Update \"OINV\" set \"U_PrintNo\"='1',\"U_PrintStatus\"='Printed',\"U_PrintDate\"=(select CURRENT_TIMESTAMP from DUMMY),\"U_PrintDTM\"=(select CURRENT_TIME from DUMMY) Where \"DocEntry\"='" + _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString() + "'";
                                        Utilities.ExecuteSQL(ref Rs, Qry);
                                    }
                                    else
                                    {
                                        ServiceExportReportToPdfFile3();
                                        string Qry = "Update \"OINV\" set \"U_PrintNo\"= \"U_PrintNo\" + 1 Where \"DocEntry\"='" + _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString() + "'";
                                        Utilities.ExecuteSQL(ref Rs, Qry);
                                    }
                                }
                                
                            }

                            finally
                            {
                                //_Form.Freeze(false);
                            }
                            #endregion
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
                            GRPOINVNo = _Form.DataSources.DBDataSources.Item("OINV").GetValue("U_GRPONum", 0).ToString();
                            if (GRPOINVNo != "")
                                _Form.Items.Item("GRPO").Enabled = false;
                            else
                                _Form.Items.Item("GRPO").Enabled = true;

                            break;

                    }
                }
                else
                {
                    switch (pVal.MenuUID)
                    {

                        case Constants.System_Menus.mnu_ADD:
                            GRPOINVNo = _Form.DataSources.DBDataSources.Item("OINV").GetValue("U_GRPONum", 0).ToString();
                            if (GRPOINVNo != "")
                                _Form.Items.Item("GRPO").Enabled = false;
                            else
                                _Form.Items.Item("GRPO").Enabled = true;

                            break;

                        case Constants.System_Menus.mnu_FIND:
                            GRPOINVNo = _Form.DataSources.DBDataSources.Item("OINV").GetValue("U_GRPONum", 0).ToString();
                            if (GRPOINVNo != "")
                                _Form.Items.Item("GRPO").Enabled = false;
                            else
                                _Form.Items.Item("GRPO").Enabled = true;

                            break;
                        case Constants.System_Menus.mnu_FIRST:
                            GRPOINVNo = _Form.DataSources.DBDataSources.Item("OINV").GetValue("U_GRPONum", 0).ToString();
                            if (GRPOINVNo != "")
                                _Form.Items.Item("GRPO").Enabled = false;
                            else
                                _Form.Items.Item("GRPO").Enabled = true;

                            break;
                        case Constants.System_Menus.mnu_LAST:
                            GRPOINVNo = _Form.DataSources.DBDataSources.Item("OINV").GetValue("U_GRPONum", 0).ToString();
                            if (GRPOINVNo != "")
                                _Form.Items.Item("GRPO").Enabled = false;
                            else
                                _Form.Items.Item("GRPO").Enabled = true;

                            break;
                        case Constants.System_Menus.mnu_NEXT:
                            GRPOINVNo = _Form.DataSources.DBDataSources.Item("OINV").GetValue("U_GRPONum", 0).ToString();
                            if (GRPOINVNo != "")
                                _Form.Items.Item("GRPO").Enabled = false;
                            else
                                _Form.Items.Item("GRPO").Enabled = true;

                            break;
                        case Constants.System_Menus.mnu_PREVIOUS:
                            GRPOINVNo = _Form.DataSources.DBDataSources.Item("OINV").GetValue("U_GRPONum", 0).ToString();
                            if (GRPOINVNo != "")
                                _Form.Items.Item("GRPO").Enabled = false;
                            else
                                _Form.Items.Item("GRPO").Enabled = true;

                            break;
                        case Constants.System_Menus.mnu_REFRESH:
                            GRPOINVNo = _Form.DataSources.DBDataSources.Item("OINV").GetValue("U_GRPONum", 0).ToString();
                            if (GRPOINVNo != "")
                                _Form.Items.Item("GRPO").Enabled = false;
                            else
                                _Form.Items.Item("GRPO").Enabled = true;

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

        #region Form Default
        internal void FormDefault()
        {
            try
            {
                _ExistItem = _Form.Items.Item("2");
                _NewItem = _Form.Items.Add("BtnPrint", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 5;
                _NewItem.Width = _ExistItem.Width + 10;
                _NewItem.Height = _ExistItem.Height;
                _NewItem.Top = _ExistItem.Top;
                _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                _Button.Caption = "Print";

                _ExistItem = _Form.Items.Item("2");
                _NewItem = _Form.Items.Add("GRPO", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 450;
                _NewItem.Width = _ExistItem.Width + 20;
                _NewItem.Height = _ExistItem.Height;
                _NewItem.Top = _ExistItem.Top;
                _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                _Button.Caption = "GRPO Creation";


            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
        #endregion

        #region Call and Show Report

        private void CallReport()
        {
            //Thread myThread = new Thread(new ThreadStart(ShowReport));
            //myThread.SetApartmentState(ApartmentState.STA);
            ////Thread myThread = new Thread(new ParameterizedThreadStart(ShowReport)); //terized Thread(new ThreadStart(ShowReport));
            //myThread.Start();
            ShowReport();
        }

        protected void ShowReport()
        {
            try
            {
                //const string _strServerName = "192.168.2.10:30015";
                //const string _strLoginName = "SYSTEM";
                //const string _strPassword = "Passw0rd";




                objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
                CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();
                ParameterFields paramFields = new ParameterFields();
                ParameterField paramField = new ParameterField();
                ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
                ParameterRangeValue rangeVal = new ParameterRangeValue();

                //  ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + sReportName;
                ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "InvoiceLayout.rpt";
                objCryRep.Load(ReportPath);

                objConInfo.ServerName = ConfigurationManager.AppSettings["DatabaseServerName"].ToString();//_strServerName;

                //objConInfo.ServerName = "192.168.2.10";

                objConInfo.DatabaseName = Utilities.Application.Company.CompanyDB;

                // objConInfo.DatabaseName = ConfigurationManager.AppSettings["Database"].ToString();// "SAPHANATEST";
                objConInfo.UserID = ConfigurationManager.AppSettings["dbUser"].ToString();//_strLoginName;
                //objConInfo.UserID = "SYSTEM";
                objConInfo.Password = ConfigurationManager.AppSettings["dbPassword"].ToString();//_strPassword;
                //objConInfo.Password = "Passw0rd";
                objConInfo.Type = ConnectionInfoType.CRQE;




                ReportDisplay Report = new ReportDisplay();
                Report.Refresh();


                foreach (CrystalDecisions.CrystalReports.Engine.Table objTable in objCryRep.Database.Tables)
                {
                    objTabLoginInfo = objTable.LogOnInfo;
                    objTabLoginInfo.ConnectionInfo = objConInfo;

                    objTable.ApplyLogOnInfo(objTabLoginInfo);
                    objTabLoginInfo.ConnectionInfo.AllowCustomConnection = true;
                }


                _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                string _ObjectId = "13";

                objCryRep.Refresh();
                objCryRep.VerifyDatabase();
               
                objCryRep.SetParameterValue("DOCKEY@", _DocEntry);
                objCryRep.SetParameterValue("OBJECTID@", _ObjectId);


                Report.crystalReportViewer1.ShowGroupTreeButton = false;
                Report.crystalReportViewer1.ReportSource = objCryRep;
                Report.crystalReportViewer1.Refresh();
                Report.crystalReportViewer1.Show();

                Report.ShowDialog();
                
                Report.crystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None;


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void CallReport1()
        {
            //Thread myThread = new Thread(new ThreadStart(ShowReport1));
            //myThread.SetApartmentState(ApartmentState.STA);
            ////Thread myThread = new Thread(new ParameterizedThreadStart(ShowReport)); //terized Thread(new ThreadStart(ShowReport));
            //myThread.Start();
            ShowReport1();
        }

        protected void ShowReport1()
        {
            try
            {
                //const string _strServerName = "192.168.2.10:30015";
                //const string _strLoginName = "SYSTEM";
                //const string _strPassword = "Passw0rd";




                objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
                CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();
                ParameterFields paramFields = new ParameterFields();
                ParameterField paramField = new ParameterField();
                ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
                ParameterRangeValue rangeVal = new ParameterRangeValue();

                //  ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + sReportName;
                ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "InvoiceLayout1.rpt";
                objCryRep.Load(ReportPath);

                objConInfo.ServerName = ConfigurationManager.AppSettings["DatabaseServerName"].ToString();//_strServerName;

                //objConInfo.ServerName = "192.168.2.10";

                objConInfo.DatabaseName = Utilities.Application.Company.CompanyDB;

                // objConInfo.DatabaseName = ConfigurationManager.AppSettings["Database"].ToString();// "SAPHANATEST";
                objConInfo.UserID = ConfigurationManager.AppSettings["dbUser"].ToString();//_strLoginName;
                //objConInfo.UserID = "SYSTEM";
                objConInfo.Password = ConfigurationManager.AppSettings["dbPassword"].ToString();//_strPassword;
                //objConInfo.Password = "Passw0rd";
                objConInfo.Type = ConnectionInfoType.CRQE;




                ReportDisplay Report = new ReportDisplay();
                Report.Refresh();


                foreach (CrystalDecisions.CrystalReports.Engine.Table objTable in objCryRep.Database.Tables)
                {
                    objTabLoginInfo = objTable.LogOnInfo;
                    objTabLoginInfo.ConnectionInfo = objConInfo;

                    objTable.ApplyLogOnInfo(objTabLoginInfo);
                    objTabLoginInfo.ConnectionInfo.AllowCustomConnection = true;
                    //bool s = objTable.TestConnectivity();
                    // string s = Utilities.Application.Company.CompanyDB + ".dbo." + objTable.Location.Substring(objTable.Location.LastIndexOf(".") + 1);
                    //objTable.Location = Utilities.Application.Company.CompanyDB + ".dbo." + objTable.Location.Substring(objTable.Location.LastIndexOf(".") + 1);

                }


                //int i = Convert.ToInt16(_DBDataH.GetValue("U_DocNum", 0));

                //objCryRep.SetParameterValue("DOCNUM", i);
                //objCryRep.SetDatabaseLogon("System", "Passw0rd", "192.168.2.10:30015","SAPHANATEST");

                //string _ExtParam = "N";
                _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                string _ObjectId = "13";

                objCryRep.Refresh();
                objCryRep.VerifyDatabase();
                //objCryRep.SetParameterValue("ExtParam@", _ExtParam);
                objCryRep.SetParameterValue("DOCKEY@", _DocEntry);
                objCryRep.SetParameterValue("OBJECTID@", _ObjectId);


                Report.crystalReportViewer1.ShowGroupTreeButton = false;
                Report.crystalReportViewer1.ReportSource = objCryRep;
                Report.crystalReportViewer1.Refresh();
                Report.crystalReportViewer1.Show();

                Report.ShowDialog();
                // Report.crystalReportViewer1.
                Report.crystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None;


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void CallReport2()
        {
            //Thread myThread = new Thread(new ThreadStart(ShowReport2));
            //myThread.SetApartmentState(ApartmentState.STA);
            ////Thread myThread = new Thread(new ParameterizedThreadStart(ShowReport)); //terized Thread(new ThreadStart(ShowReport));
            //myThread.Start();
            ShowReport2();
        }

        protected void ShowReport2()
        {
            try
            {
                //const string _strServerName = "192.168.2.10:30015";
                //const string _strLoginName = "SYSTEM";
                //const string _strPassword = "Passw0rd";




                objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
                CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();
                ParameterFields paramFields = new ParameterFields();
                ParameterField paramField = new ParameterField();
                ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
                ParameterRangeValue rangeVal = new ParameterRangeValue();

                //  ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + sReportName;
                ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "InvoiceLayout2.rpt";
                objCryRep.Load(ReportPath);

                objConInfo.ServerName = ConfigurationManager.AppSettings["DatabaseServerName"].ToString();//_strServerName;

                //objConInfo.ServerName = "192.168.2.10";

                objConInfo.DatabaseName = Utilities.Application.Company.CompanyDB;

                // objConInfo.DatabaseName = ConfigurationManager.AppSettings["Database"].ToString();// "SAPHANATEST";
                objConInfo.UserID = ConfigurationManager.AppSettings["dbUser"].ToString();//_strLoginName;
                //objConInfo.UserID = "SYSTEM";
                objConInfo.Password = ConfigurationManager.AppSettings["dbPassword"].ToString();//_strPassword;
                //objConInfo.Password = "Passw0rd";
                objConInfo.Type = ConnectionInfoType.CRQE;




                ReportDisplay Report = new ReportDisplay();
                Report.Refresh();


                foreach (CrystalDecisions.CrystalReports.Engine.Table objTable in objCryRep.Database.Tables)
                {
                    objTabLoginInfo = objTable.LogOnInfo;
                    objTabLoginInfo.ConnectionInfo = objConInfo;

                    objTable.ApplyLogOnInfo(objTabLoginInfo);
                    objTabLoginInfo.ConnectionInfo.AllowCustomConnection = true;
                    //bool s = objTable.TestConnectivity();
                    // string s = Utilities.Application.Company.CompanyDB + ".dbo." + objTable.Location.Substring(objTable.Location.LastIndexOf(".") + 1);
                    //objTable.Location = Utilities.Application.Company.CompanyDB + ".dbo." + objTable.Location.Substring(objTable.Location.LastIndexOf(".") + 1);

                }


                //int i = Convert.ToInt16(_DBDataH.GetValue("U_DocNum", 0));

                //objCryRep.SetParameterValue("DOCNUM", i);
                //objCryRep.SetDatabaseLogon("System", "Passw0rd", "192.168.2.10:30015","SAPHANATEST");

                //string _ExtParam = "N";
                _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                string _ObjectId = "13";

                objCryRep.Refresh();
                objCryRep.VerifyDatabase();
                //objCryRep.SetParameterValue("ExtParam@", _ExtParam);
                objCryRep.SetParameterValue("DOCKEY@", _DocEntry);
                objCryRep.SetParameterValue("OBJECTID@", _ObjectId);


                Report.crystalReportViewer1.ShowGroupTreeButton = false;
                Report.crystalReportViewer1.ReportSource = objCryRep;
                Report.crystalReportViewer1.Refresh();
                Report.crystalReportViewer1.Show();

                Report.ShowDialog();
                // Report.crystalReportViewer1.
                Report.crystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None;


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        #region EXPORT REPORT TO PDF FILE

        /// <summary>
        /// Export Report into pdf document
        /// </summary>
        /// <param name="oCustomerCode"></param>
        /// <param name="oContactPersonName"></param>
        /// <param name="oMail"></param>
        /// <param name="_LayoutName"></param>
        /// <param name="oPrinter"></param>
        /// <param name="oLayoutID"></param>
        /// <param name="DocNum"></param>
        /// <param name="DocEntry"></param>
        /// <param name="oContactPersonCode"></param>
        /// <param name="_DocType"></param>
        private void ExportReportToPdfFile()
        {
            DataSet ds = LoadCompanySetting();
            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();
            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;
            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();
            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();
            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();
            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "InvoiceLayout.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 25);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }
            }
        }
        private void ExportReportToPdfFile1()
        {
            DataSet ds = LoadCompanySetting();

            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();

            //string DatabaseServerName = ConfigurationManager.AppSettings["DatabaseServerName"].ToString();//_strServerName;
            //string dbUser = ConfigurationManager.AppSettings["dbUser"].ToString();//_strLoginName;
            //string dbPassword = ConfigurationManager.AppSettings["dbPassword"].ToString();//_strPassword;
            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;



            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "InvoiceLayout1.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 25);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }

            }
        }
        private void ExportReportToPdfFile2()
        {
            DataSet ds = LoadCompanySetting();

            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();

            //string DatabaseServerName = ConfigurationManager.AppSettings["DatabaseServerName"].ToString();//_strServerName;
            //string dbUser = ConfigurationManager.AppSettings["dbUser"].ToString();//_strLoginName;
            //string dbPassword = ConfigurationManager.AppSettings["dbPassword"].ToString();//_strPassword;
            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;



            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "InvoiceLayout2.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 25);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }

            }
        }
        private void ExportReportToPdfFile3()
        {
            DataSet ds = LoadCompanySetting();

            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();

            //string DatabaseServerName = ConfigurationManager.AppSettings["DatabaseServerName"].ToString();//_strServerName;
            //string dbUser = ConfigurationManager.AppSettings["dbUser"].ToString();//_strLoginName;
            //string dbPassword = ConfigurationManager.AppSettings["dbPassword"].ToString();//_strPassword;
            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;

            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "InvoiceLayout3.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 1);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }
            }
        }
        private void FAExportReportToPdfFile()
        {
            DataSet ds = LoadCompanySetting();

            //string DatabaseServerName = ConfigurationManager.AppSettings["DatabaseServerName"].ToString();//_strServerName;
            //string dbUser = ConfigurationManager.AppSettings["dbUser"].ToString();//_strLoginName;
            //string dbPassword = ConfigurationManager.AppSettings["dbPassword"].ToString();//_strPassword;

            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();

            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;

            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "FAInvoiceLayout.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 25);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }
            }
        }
        private void FAExportReportToPdfFile1()
        {
            DataSet ds = LoadCompanySetting();

            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();

            //string DatabaseServerName = ConfigurationManager.AppSettings["DatabaseServerName"].ToString();//_strServerName;
            //string dbUser = ConfigurationManager.AppSettings["dbUser"].ToString();//_strLoginName;
            //string dbPassword = ConfigurationManager.AppSettings["dbPassword"].ToString();//_strPassword;
            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;


            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "FAInvoiceLayout1.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 25);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }
            }
        }
        private void FAExportReportToPdfFile2()
        {
            DataSet ds = LoadCompanySetting();

            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();

            //string DatabaseServerName = ConfigurationManager.AppSettings["DatabaseServerName"].ToString();//_strServerName;
            //string dbUser = ConfigurationManager.AppSettings["dbUser"].ToString();//_strLoginName;
            //string dbPassword = ConfigurationManager.AppSettings["dbPassword"].ToString();//_strPassword;
            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;



            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "FAInvoiceLayout2.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 25);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }
            }
        }
        private void FAExportReportToPdfFile3()
        {
            DataSet ds = LoadCompanySetting();

            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();

            //string DatabaseServerName = ConfigurationManager.AppSettings["DatabaseServerName"].ToString();//_strServerName;
            //string dbUser = ConfigurationManager.AppSettings["dbUser"].ToString();//_strLoginName;
            //string dbPassword = ConfigurationManager.AppSettings["dbPassword"].ToString();//_strPassword;
            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;

            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "FAInvoiceLayout3.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 25);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }
            }
        }
        private void ServiceExportReportToPdfFile()
        {
            DataSet ds = LoadCompanySetting();
            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();

            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;

            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "InvoiceLayout_1.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 25);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }
            }
        }
        private void ServiceExportReportToPdfFile1()
        {
            DataSet ds = LoadCompanySetting();

            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();

            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;


            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "InvoiceLayout1_1.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 25);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }
            }
        }
        private void ServiceExportReportToPdfFile2()
        {
            DataSet ds = LoadCompanySetting();

            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();

            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;

            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "InvoiceLayout2_1.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 25);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }
            }
        }
        private void ServiceExportReportToPdfFile3()
        {
            DataSet ds = LoadCompanySetting();

            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();

            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;

            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "InvoiceLayout3_1.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    _DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 25);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }
            }
        }

        #endregion

        #region PrintLayout

        private void PrintLayout()
        {
            ReportDocument cryRpt = new ReportDocument();
            ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "InvoiceLayout.rpt";
            cryRpt.Load(ReportPath);

            TableLogOnInfos crtableLogoninfos = new TableLogOnInfos();
            TableLogOnInfo crtableLogoninfo = new TableLogOnInfo();
            ConnectionInfo crConnectionInfo = new ConnectionInfo();
            Tables CrTables;

            crConnectionInfo.ServerName = ConfigurationManager.AppSettings["DatabaseServerName"].ToString();//_strServerName;
            crConnectionInfo.DatabaseName = Utilities.Application.Company.CompanyDB;
            crConnectionInfo.UserID = ConfigurationManager.AppSettings["dbUser"].ToString();//_strLoginName;
            crConnectionInfo.Password = ConfigurationManager.AppSettings["dbPassword"].ToString();//_strPassword;

            CrTables = cryRpt.Database.Tables;
            foreach (CrystalDecisions.CrystalReports.Engine.Table CrTable in CrTables)
            {
                crtableLogoninfo = CrTable.LogOnInfo;
                crtableLogoninfo.ConnectionInfo = crConnectionInfo;
                CrTable.ApplyLogOnInfo(crtableLogoninfo);
            }

            cryRpt.Refresh();
            cryRpt.PrintToPrinter(2, true, 1, 2);
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

        #region Upload data of invoice on IRD Portal

        public void IRD(string DocNum)
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
                              " OINV.\"DocNum\",OINV.\"U_LD\",OCRD.\"LicTradNum\",OINV.\"DocEntry\" from OINV inner join OCRD ON OCRD.\"CardCode\" = OINV.\"CardCode\" " +
                              " INNER JOIN NNM1 ON OINV.\"Series\" = NNM1.\"Series\" AND OINV.\"ObjType\" = NNM1.\"ObjectCode\" where \"DocDate\" > '17.07.2022' AND OINV.\"DocStatus\" = 'O' and \"DocNum\" = '" + DocNum + "' Order by OINV.\"DocEntry\" desc";
                Utilities.ExecuteSQL(ref _RS, oSql);
                string aa = _RS.Fields.Item("U_LD").Value.ToString();
                if (_RS.RecordCount > 0)
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    BillViewModel p = new BillViewModel
                    {
                        username = WebUserName.ToString(),
                        password = WebPassword.ToString(),
                        seller_pan = seller_pan.ToString(),
                        buyer_pan = _RS.Fields.Item("LicTradNum").Value.ToString().Trim(),
                        buyer_name = _RS.Fields.Item("CardName").Value.ToString().Trim(),
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
                    client.BaseAddress = new Uri(InvBaseAddress.ToString());
                    var response = client.PostAsJsonAsync(Invresponse.ToString(), p).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        var responseCode = response.Content.ReadAsStringAsync();
                        if (responseCode.Result == "200")
                        {
                            SAPbobsCOM.Recordset _RS12 = null;
                            string _Qry = "Update \"OINV\" set \"U_SyncIRD\"='Y',\"U_IsRealTm\" = 'Y',\"U_SyncMsgID\" = '200' Where \"DocNum\"='" + DocNum + "' and \"DocEntry\" = '" + _RS.Fields.Item("DocEntry").Value.ToString().Trim() + "'";
                            Utilities.ExecuteSQL(ref _RS12, _Qry);

                            Console.Write(responseCode.Result);//responseCode 200 means successful
                            Console.ReadLine();
                        }
                        else
                        {
                            SAPbobsCOM.Recordset _RS13 = null;
                            string _Qry1 = "Update \"OINV\" set \"U_SyncIRD\"='N',\"U_IsRealTm\" = 'N',\"U_SyncMsgID\" = '" + responseCode.Result + "' Where \"DocNum\"='" + DocNum + "' and \"DocEntry\" = '" + _RS.Fields.Item("DocEntry").Value.ToString().Trim() + "'";
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

        #region SendEmail
        public void SendEmail(string DocNum)
        {
            string PDFFilePathEmails = string.Empty;
            string EmailBody = string.Empty;
            string EmailSubject = string.Empty;
            DataSet ds = LoadCompanySetting();
               SAPbobsCOM.Recordset _RS = null;
               string oSql = "select OINV.\"DocEntry\",OCRD.\"E_Mail\" from OINV inner join OCRD ON OCRD.\"CardCode\" = OINV.\"CardCode\" where OINV.\"DocNum\" = '" + DocNum + "'";
                Utilities.ExecuteSQL(ref _RS, oSql);
                if (_RS.RecordCount > 0)
                {
                    ExportReportToPdfFileForEmail(_RS.Fields.Item("DocEntry").Value.ToString());
                    string FileName = _RS.Fields.Item("DocEntry").Value.ToString() + ".pdf";
                    PDFFilePathEmails = (ds.Tables[0].Rows[0]["SavePDFPath"].ToString() + "\\" + FileName);
                    //EmailBody = ds.Tables[0].Rows[0]["EmailBody"].ToString();
                    EmailBody = "Dear Sir, \nPlease find the enclosed copy of Tax Invoice for your reference. The attached invoice memo is only a copy of the TAX INVOICE issued in favor of you for your easy reference and better communication. If you do not receive the TAX invoice, please contact to respective contact details mentioned below: \n\r For, Pokhara Branch \n Email id: accounts.sujaldairy@sujal.com.np \r\n Phone no: 061-520460,061-520355, 9801240245\n\r For, Kathmandu Branch \nEmail id: account.sujaldairyb@sujal.com.np " + Environment.NewLine + " Phone no: 01-4316344, 9801240246 \n\r Thanks & Regards \n\r Sujal Dairy Pvt. Ltd.";
                    EmailSubject = ds.Tables[0].Rows[0]["EmailSubject"].ToString();

                    if (_RS.Fields.Item("E_Mail").Value.ToString() != "")
                        SendMail(_RS.Fields.Item("E_Mail").Value.ToString(), "Invoice Attachment", EmailBody, PDFFilePathEmails.ToString());
                    Utilities.ShowSucessMessage("PDF generation and Send emails operation completed Sucessfully");
                }
                else
                {
                    Utilities.ShowErrorMessage("No record found");

                }
        }

        #endregion

        #region ExportReportToPdfFileForEmail

        private void ExportReportToPdfFileForEmail(string DocEntry)
        {
            DataSet ds = LoadCompanySetting();
            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();
            string SavePDFPath = ds.Tables[0].Rows[0]["SavePDFPath"].ToString().Trim();
            string Database = Utilities.Application.Company.CompanyDB;
            //string currentUser = Utilities.Application.Company.UserName;

            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;

                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "EmailInvoiceCopy.rpt";
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    objCryRep.SetDatabaseLogon(dbUser.ToString(), dbPassword.ToString(), DatabaseServerName.ToString(), Utilities.Application.Company.CompanyDB);
                    objConInfo.ServerName = DatabaseServerName.ToString();
                    objConInfo.DatabaseName = Database.ToString();
                    objConInfo.UserID = dbUser.ToString();
                    objConInfo.Password = dbPassword.ToString();
                    int n = objCryRep.Database.Tables.Count;
                    // Set the logon credentials for all tables
                    _DocEntry = DocEntry;// _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();

                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);

                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                   // objCryRep.SetParameterValue("USERNAME@", currentUser);

                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************

                    // Export Report To Define Application Path --User
                    string PDFFilePath = SavePDFPath.ToString();
                    objCryRep.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, PDFFilePath + "\\" + _ImportFileName);
                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }
            }
        }

        #endregion

        #region SendMail

        public void SendMail(string ToEmail, string Subject, string EmailBody, string AttachmentFilePath)
        {
            try
            {

                DataSet ds = LoadCompanySetting();

                string FromEmail = ds.Tables[0].Rows[0]["FromEmail"].ToString().Trim();
                string FromPassword = ds.Tables[0].Rows[0]["FromEmailPassword"].ToString().Trim();

                MailMessage mail = new MailMessage();
                //SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                SmtpClient SmtpServer = new SmtpClient("mail.sujal.com.np");
                mail.From = new MailAddress(FromEmail.ToString().Trim());
                mail.To.Add(ToEmail.ToString());
                mail.Subject = Subject.ToString();
                mail.Body = EmailBody.ToString();

                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(AttachmentFilePath.ToString());
                mail.Attachments.Add(attachment);

                //SmtpServer.Port = 587;
                SmtpServer.Port = 587;//465
                SmtpServer.Credentials = new System.Net.NetworkCredential(FromEmail.ToString().Trim(), FromPassword.ToString().Trim());
                SmtpServer.EnableSsl = true;
                ServicePointManager.ServerCertificateValidationCallback = delegate(object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
             
                SmtpServer.Send(mail);

                attachment.Dispose();
                SmtpServer = null;
                mail.Dispose();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Cursor.Current = System.Windows.Forms.Cursors.Default;
            }

        }

        #endregion

        #region CreateGRPO

        private void CreateGRPO(int _DocNum, int _DocEntry, string BPCode, string BranchCode, string PostingDate, string DeliveryDate, string DocumentDate ,string strseries,string WhseCode)
        {
            try
            {
                string lineInTrn = "";
                string MainTbl = "";
                string LineTbl = "";
                int skipLine = 0;

               
                SAPbobsCOM.BoObjectTypes objType = SAPbobsCOM.BoObjectTypes.oInvoices;
                SAPbobsCOM.Documents document;

                MainTbl = "OINV";
                LineTbl = "INV1";
                objType = SAPbobsCOM.BoObjectTypes.oInvoices;
                Utilities.ShowWarningMessage("Please wait...");

                string sql = "  SELECT T3.\"Substitute\",T.\"DiscPrcnt\",T1.\"WhsCode\",T1.\"Price\", T1.\"ItemCode\",T1.\"Dscription\",T1.\"Quantity\",T1.\"LineNum\",T2.\"ManSerNum\",T2.\"ManBtchNum\",T1.\"DocEntry\" "
                             + String.Format(" FROM {0} T1 INNER JOIN {1} T ON T.\"DocEntry\" =T1.\"DocEntry\" ", LineTbl, MainTbl)
                             + " Inner join OITM T2 on T1.\"ItemCode\" =T2.\"ItemCode\" "
                             +"left outer join OSCN T3 ON T1.\"ItemCode\" = T3.\"ItemCode\" AND T3.\"CardCode\" = '"+ BPCode +"' "
                             + String.Format(" where T.\"DocEntry\" ={0} ", _DocEntry)
                             + " group by T.\"DiscPrcnt\",T1.\"WhsCode\",T1.\"Price\", T1.\"ItemCode\",T1.\"Dscription\",T1.\"Quantity\",T1.\"LineNum\",T2.\"ManSerNum\",T2.\"ManBtchNum\",T1.\"DocEntry\",T3.\"Substitute\""
                             + " order by \"LineNum\" asc ";

                Utilities.ExecuteSQL(ref _RSGRPO, sql);
               
                //return Isuccess;
                /*End DIAPI Company Connection*/
                /*Start GRPO Creation by DIAPI based on Delivery*/
                SAPbobsCOM.Documents _GRPO = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
                document = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(objType);
                document.GetByKey(_DocEntry);

                _GRPO.CardCode = BPCode.ToString();
                //_GRPO.DocDate = Utilities.ToDate(DateTime.Now.ToString("yyyyMMdd"));

                DateTime parsedPostingDate = DateTime.Parse(PostingDate);
                parsedPostingDate.ToString("yyyyMMdd");

                DateTime parsedPDeliveryDate = DateTime.Parse(DeliveryDate);
                parsedPDeliveryDate.ToString("yyyyMMdd");

                DateTime parsedPDocumentDate = DateTime.Parse(DeliveryDate);
                parsedPDocumentDate.ToString("yyyyMMdd");


                _GRPO.DocDate = Utilities.ToDate(parsedPostingDate.ToString("yyyyMMdd"));
                _GRPO.DocDueDate = Utilities.ToDate(parsedPDeliveryDate.ToString("yyyyMMdd"));
                _GRPO.TaxDate = Utilities.ToDate(parsedPDocumentDate.ToString("yyyyMMdd"));
              //  _GRPO.Series = Convert.ToInt16(strseries.ToString());
                _GRPO.BPL_IDAssignedToInvoice = Convert.ToInt16(BranchCode.ToString());
                _GRPO.UserFields.Fields.Item("U_InvNo").Value = _DocEntry.ToString();

                if (_RSGRPO.RecordCount > 0)
                {

                    lineInTrn = "";
                    //DataRow[] rows = dt.Select("DocEntry ='" + dt.Rows[0]["DocEntry"].ToString() + "'");
                    _GRPO.DiscountPercent = double.Parse(_RSGRPO.Fields.Item("DiscPrcnt").Value.ToString());
                    if (_RSGRPO.RecordCount > 0)
                    {

                        for (int k = 0; k <= _RSGRPO.RecordCount - 1; k++)
                        {
                            _GRPO.Lines.SetCurrentLine(k);
                            if (_RSGRPO.Fields.Item("ManSerNum").Value.ToString() == "Y")
                            {
                                SAPbobsCOM.Recordset Rs = null;

                                string Qry = "SELECT T3.\"ItemCode\",T3.\"DistNumber\",T3.\"MnfSerial\",T3.\"U_KeyNo\",T3.\"U_ModelYear\",T3.\"U_RegNo\" FROM ODLN T0 INNER JOIN DLN1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                                             " INNER JOIN SRI1 T2 on T1.\"ItemCode\"=T2.\"ItemCode\" and T1.\"LineNum\" = T2.\"BaseLinNum\" and (T1.\"DocEntry\"=T2.\"BaseEntry\" and T1.\"ObjType\"=T2.\"BaseType\") " +
                                             " INNER JOIN OSRN T3 on T3.\"ItemCode\"=T2.\"ItemCode\" and T2.\"SysSerial\"=T3.\"SysNumber\" where  T0.\"DocEntry\" = '" + _DocEntry + "' and T3.\"ItemCode\" = '" + _RSGRPO.Fields.Item("itemcode").Value.ToString() + "' AND T1.\"LineNum\" = " + k + " ";
                                Utilities.ExecuteSQL(ref Rs, Qry);

                                for (int j = 0; j <= Rs.RecordCount - 1; j++)
                                {
                                    if (document.Lines.SerialNumbers.InternalSerialNumber != "")
                                    {
                                        _GRPO.Lines.SerialNumbers.SetCurrentLine(j);
                                        _GRPO.Lines.SerialNumbers.ManufacturerSerialNumber = Rs.Fields.Item("MnfSerial").Value.ToString();
                                        _GRPO.Lines.SerialNumbers.InternalSerialNumber = Rs.Fields.Item("DistNumber").Value.ToString();
                                        _GRPO.Lines.SerialNumbers.Add();

                                    }
                                    Rs.MoveNext();
                                }
                            }
                            _GRPO.Lines.ItemCode = _RSGRPO.Fields.Item("Substitute").Value.ToString();//_RSGRPO.Fields.Item("itemcode").Value.ToString();
                            _GRPO.Lines.WarehouseCode = WhseCode.ToString();
                            _GRPO.Lines.Quantity = double.Parse(_RSGRPO.Fields.Item("Quantity").Value.ToString());
                            _GRPO.Lines.UnitPrice = double.Parse(_RSGRPO.Fields.Item("Price").Value.ToString());

                            _GRPO.Lines.Add();
                            _RSGRPO.MoveNext();
                        }

                        int RetVal = _GRPO.Add();
                        if (RetVal == 0)
                        {
                            string objCode = "";
                            Utilities.Application.Company.GetNewObjectCode(out objCode);
                            SAPbobsCOM.Recordset _Rs = null;
                            _Rs1 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string StrQry1 = "select \"DocNum\" from OINV where \"DocEntry\" = '" + objCode + "'";
                            Utilities.ExecuteSQL(ref _Rs1, StrQry1);

                            string StrQry = "update OINV set \"U_GRPONum\" = '" + objCode + "' where \"DocEntry\" = '" + _DocEntry + "'";
                            Utilities.ExecuteSQL(ref _Rs, StrQry);
                            WriteLogFile("Log", "GRPO Created Sucessfully", ":  GRPO Entry :" + objCode);
                           // System.Runtime.InteropServices.Marshal.ReleaseComObject(Utilities.Application.Company);//For releasing the company object
                        }
                        else
                        {
                            _Form.Items.Item("GRPO").Enabled = true;
                            string GetLastErrorDescription = Utilities.Application.Company.GetLastErrorDescription();
                            WriteLogFile("Log", "GRPO Creation", "Database Name : " + GetLastErrorDescription);
                           // System.Runtime.InteropServices.Marshal.ReleaseComObject(Utilities.Application.Company);
                            throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }
                    }
                }

                Utilities.ShowWarningMessage("Operation completed successfully.");
            }
            catch (Exception ex)
            {
                Utilities.LogEntry("Error: "
           + ex.Message, "Error in GRPO : DocEntry "
           + _DocNum + " @ " + System.DateTime.Now.ToString());
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

    }
}
                                     