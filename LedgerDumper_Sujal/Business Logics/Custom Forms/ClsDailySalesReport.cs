using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace InvoicePrinting_Sujal
{
    class ClsDailySalesReport : Base
    {
        string _Sql = "";
        SAPbobsCOM.Recordset _RS;
        SAPbouiCOM.EditText _Edit;
        SAPbouiCOM.ComboBox _Combo;
        private SAPbouiCOM.Grid _Grid = null;
        SAPbouiCOM.DataTable _DataTable;
        ChooseFromListEvent _SysCFLEvent;
        Conditions _Cons;
        Condition _Con;
        private SAPbouiCOM.CommonSetting _CommonSetting;
        private SAPbouiCOM.RowHeaders _RowHeaders;
        string ItemCode = "";
        SAPbouiCOM.ComboBox _ComboBox;

        #region ITEM EVENT
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            #region BEFORE ACTION ACTION TRUE
            if (pVal.Before_Action == true)
            {
                switch (pVal.EventType)
                {
                }
            }
            #endregion

            #region BEFORE AFTER ACTION TRUE
            else if (pVal.Before_Action == false)
            {

                switch (pVal.EventType)
                {
                    #region ITEM PRESS EVENT
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "7" && pVal.Action_Success == true)
                        {
                            string FromDt = _Form.DataSources.UserDataSources.Item("FromDt").ValueEx.ToString();
                            string ToDt = _Form.DataSources.UserDataSources.Item("ToDt").ValueEx.ToString();
                            string _Branch = _Form.DataSources.UserDataSources.Item("UsrBranch").ValueEx.ToString();
                            this.OpenGrid(FromDt, ToDt, _Branch);
                        }
                        break;
                    #endregion
                }
            }
            #endregion
        }
        #endregion


        #region Open Grid

        internal void OpenGrid(string FromDt, string ToDt,string _Branch)
        {
            try
            {
                this.Form.Visible = false;
                string SQL = "CALL SALESREGISTER ('" + FromDt + "','" + ToDt + "','" + _Branch + "')";
                this._Object = new ClsOpenSalesReportGrid();
                this.OpenChildForm(Constants.Forms.Grid_Daily_SalesReport, true);
                ((ClsOpenSalesReportGrid)_Object).SalesReport(SQL, FromDt, ToDt);
            }
            finally
            {
               // this.Form.Visible = true;
            }


        }

        #endregion

        #region Form Default
        internal void FormDefault()
        {
            

            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("9").Specific;
            Utilities.FillCombo(ref _ComboBox, "select \"BPLId\",\"BPLName\" from OBPL");
      
        }
        #endregion
    }
}
