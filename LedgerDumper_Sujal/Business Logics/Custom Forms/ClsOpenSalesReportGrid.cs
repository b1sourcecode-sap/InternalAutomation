using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;

namespace InvoicePrinting_Sujal
{
    class ClsOpenSalesReportGrid : Base
    {

        string _Sql = "";
        SAPbobsCOM.Recordset _RS, _RS1;
        SAPbouiCOM.EditText _Edit;
        SAPbouiCOM.ComboBox _Combo;
        private SAPbouiCOM.Grid _Grid = null;
        SAPbouiCOM.DataTable _DataTable1;
        ChooseFromListEvent _SysCFLEvent;
        Conditions _Cons;
        Condition _Con;
        private SAPbouiCOM.CommonSetting _CommonSetting;
        private SAPbouiCOM.RowHeaders _RowHeaders;
        string _SQL = "";
        private SAPbouiCOM.EditTextColumn _Column;

        string ItemCode = "";

        internal void SalesReport(string Qry,string FromDate,string ToDate)
        {
            _SQL = "SELECT DAYOFMONTH(TO_NVARCHAR('" + FromDate + "', 'YYYYMMDD')) as \"Day\",DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',1), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',2), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',3), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',4), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',5), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',6), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',7), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',8), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',9), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',10), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',11), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',12), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',13), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',14), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',15), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',16), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',17), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',18), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',19), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',20), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',21), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',22), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',23), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',24), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',25), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',26), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',27), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',28), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',29), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',30), 'YYYYMMDD')),DAYOFMONTH(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',31), 'YYYYMMDD')) from dummy";
            Utilities.ExecuteSQL(ref _RS, _SQL);
            

            _SQL = "SELECT MONTHNAME(TO_NVARCHAR('" + FromDate + "', 'YYYYMMDD')) as \"Month\",MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',1), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',2), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',3), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',4), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',5), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',6), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',7), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',8), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',9), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',10), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',11), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',12), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',13), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',14), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',15), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',16), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',17), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',18), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',19), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',20), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',21), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',22), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',23), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',24), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',25), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',26), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',27), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',28), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',29), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',30), 'YYYYMMDD')),MONTHNAME(TO_NVARCHAR(ADD_DAYS('" + FromDate + "',31), 'YYYYMMDD')) from dummy";
            Utilities.ExecuteSQL(ref _RS1, _SQL);
           
            string DaysMonth = "" + _RS.Fields.Item(0).Value.ToString() + " - " + _RS1.Fields.Item(0).Value.ToString() + "";

            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("ItmGrid").Specific;
            _Grid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            _Form.DataSources.DataTables.Add("dtSalesDt");
            _DataTable1 = _Form.DataSources.DataTables.Item("dtSalesDt");
            _DataTable1.ExecuteQuery(Qry);
            if (_DataTable1.Rows.Count > 0)
            {
                _Grid.DataTable = _DataTable1;
                _Grid.SelectionMode = BoMatrixSelect.ms_Single;
                //_CommonSetting = _Grid.CommonSetting;
                _Form.Items.Item("ItmGrid").Enabled = false;
                _RowHeaders = _Grid.RowHeaders;
                _RowHeaders.TitleObject.Caption = "#";
                for (int i = 0; i < _Grid.Rows.Count; i++)
                {
                    _RowHeaders.SetText(i, (i + 1).ToString());
                }
                _Grid.AutoResizeColumns();
            }
            _Column = (SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(0);
            _Column.LinkedObjectType = "4";
            _Grid.Columns.Item("1").TitleObject.Caption = "" + _RS.Fields.Item(0).Value.ToString() + " - " + _RS1.Fields.Item(0).Value.ToString() + "";
            _Grid.Columns.Item("2").TitleObject.Caption = "" + _RS.Fields.Item(1).Value.ToString() + " - " + _RS1.Fields.Item(1).Value.ToString() + "";
            _Grid.Columns.Item("3").TitleObject.Caption = "" + _RS.Fields.Item(2).Value.ToString() + " - " + _RS1.Fields.Item(2).Value.ToString() + "";
            _Grid.Columns.Item("4").TitleObject.Caption = "" + _RS.Fields.Item(3).Value.ToString() + " - " + _RS1.Fields.Item(3).Value.ToString() + "";
            _Grid.Columns.Item("5").TitleObject.Caption = "" + _RS.Fields.Item(4).Value.ToString() + " - " + _RS1.Fields.Item(4).Value.ToString() + "";
            _Grid.Columns.Item("6").TitleObject.Caption = "" + _RS.Fields.Item(5).Value.ToString() + " - " + _RS1.Fields.Item(5).Value.ToString() + "";
            _Grid.Columns.Item("7").TitleObject.Caption = "" + _RS.Fields.Item(6).Value.ToString() + " - " + _RS1.Fields.Item(6).Value.ToString() + "";
            _Grid.Columns.Item("8").TitleObject.Caption = "" + _RS.Fields.Item(7).Value.ToString() + " - " + _RS1.Fields.Item(7).Value.ToString() + "";
            _Grid.Columns.Item("9").TitleObject.Caption = "" + _RS.Fields.Item(8).Value.ToString() + " - " + _RS1.Fields.Item(8).Value.ToString() + "";
            _Grid.Columns.Item("10").TitleObject.Caption = "" + _RS.Fields.Item(9).Value.ToString() + " - " + _RS1.Fields.Item(9).Value.ToString() + "";
            _Grid.Columns.Item("11").TitleObject.Caption = "" + _RS.Fields.Item(10).Value.ToString() + " - " + _RS1.Fields.Item(10).Value.ToString() + "";
            _Grid.Columns.Item("12").TitleObject.Caption = "" + _RS.Fields.Item(11).Value.ToString() + " - " + _RS1.Fields.Item(11).Value.ToString() + "";
            _Grid.Columns.Item("13").TitleObject.Caption = "" + _RS.Fields.Item(12).Value.ToString() + " - " + _RS1.Fields.Item(12).Value.ToString() + "";
            _Grid.Columns.Item("14").TitleObject.Caption = "" + _RS.Fields.Item(13).Value.ToString() + " - " + _RS1.Fields.Item(13).Value.ToString() + "";
            _Grid.Columns.Item("15").TitleObject.Caption = "" + _RS.Fields.Item(14).Value.ToString() + " - " + _RS1.Fields.Item(14).Value.ToString() + "";
            _Grid.Columns.Item("16").TitleObject.Caption = "" + _RS.Fields.Item(15).Value.ToString() + " - " + _RS1.Fields.Item(15).Value.ToString() + "";
            _Grid.Columns.Item("17").TitleObject.Caption = "" + _RS.Fields.Item(16).Value.ToString() + " - " + _RS1.Fields.Item(16).Value.ToString() + "";
            _Grid.Columns.Item("18").TitleObject.Caption = "" + _RS.Fields.Item(17).Value.ToString() + " - " + _RS1.Fields.Item(17).Value.ToString() + "";
            _Grid.Columns.Item("19").TitleObject.Caption = "" + _RS.Fields.Item(18).Value.ToString() + " - " + _RS1.Fields.Item(18).Value.ToString() + "";
            _Grid.Columns.Item("20").TitleObject.Caption = "" + _RS.Fields.Item(19).Value.ToString() + " - " + _RS1.Fields.Item(19).Value.ToString() + "";
            _Grid.Columns.Item("21").TitleObject.Caption = "" + _RS.Fields.Item(20).Value.ToString() + " - " + _RS1.Fields.Item(20).Value.ToString() + "";
            _Grid.Columns.Item("22").TitleObject.Caption = "" + _RS.Fields.Item(21).Value.ToString() + " - " + _RS1.Fields.Item(21).Value.ToString() + "";
            _Grid.Columns.Item("23").TitleObject.Caption = "" + _RS.Fields.Item(22).Value.ToString() + " - " + _RS1.Fields.Item(22).Value.ToString() + "";
            _Grid.Columns.Item("24").TitleObject.Caption = "" + _RS.Fields.Item(23).Value.ToString() + " - " + _RS1.Fields.Item(23).Value.ToString() + "";
            _Grid.Columns.Item("25").TitleObject.Caption = "" + _RS.Fields.Item(24).Value.ToString() + " - " + _RS1.Fields.Item(24).Value.ToString() + "";
            _Grid.Columns.Item("26").TitleObject.Caption = "" + _RS.Fields.Item(25).Value.ToString() + " - " + _RS1.Fields.Item(25).Value.ToString() + "";
            _Grid.Columns.Item("27").TitleObject.Caption = "" + _RS.Fields.Item(26).Value.ToString() + " - " + _RS1.Fields.Item(26).Value.ToString() + "";
            _Grid.Columns.Item("28").TitleObject.Caption = "" + _RS.Fields.Item(27).Value.ToString() + " - " + _RS1.Fields.Item(27).Value.ToString() + "";
            _Grid.Columns.Item("29").TitleObject.Caption = "" + _RS.Fields.Item(28).Value.ToString() + " - " + _RS1.Fields.Item(28).Value.ToString() + "";
            _Grid.Columns.Item("30").TitleObject.Caption = "" + _RS.Fields.Item(29).Value.ToString() + " - " + _RS1.Fields.Item(29).Value.ToString() + "";
            _Grid.Columns.Item("31").TitleObject.Caption = "" + _RS.Fields.Item(30).Value.ToString() + " - " + _RS1.Fields.Item(30).Value.ToString() + "";
            
        }



    }
}
