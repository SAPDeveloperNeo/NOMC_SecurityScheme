using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SecurityScheme
{
    class SecondarySchemeReport
    {

        #region Variable

        private SAPbouiCOM.Form oForm, oForm1, oFormCFL;
        private SAPbouiCOM.Button oBtn = null;
        private SAPbouiCOM.Item oItem, oItem1, oItem2, oItem3;

        private SAPbouiCOM.Grid oGrid;
       



        private SAPbouiCOM.Matrix oMatrix, oMat;
        private Boolean ACTION = false;
        private SAPbobsCOM.Recordset oRecordSet, oRec1, oRec, oRecordSetforGrid;
        private int Mode;
        private int i, DelLine;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;
        private SAPbouiCOM.ChooseFromList oCFL = null;

        #endregion

        #region MenuEvent
        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId, string Type)
        {
            bool bevent = true;
            try
            {

                oForm = clsMain.SBO_Application.Forms.Item(FormId);
                if (Type == "Add")
                {
                }

                return bevent;
            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Menu Event : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
        }
        #endregion

        #region  Item event
        public bool Itemevent(ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string FormId)
        {

            try
            {
                oForm = clsMain.SBO_Application.Forms.Item(FormId);
                switch (pVal.EventType)
                {
                    //cfl
                    #region ITEM_PRESSED
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        if (pVal.BeforeAction)
                        {
                            try
                            {
                                if (pVal.ItemUID == "1" && (pVal.FormMode == 3 || pVal.FormMode == 2))
                                {
                                    Mode = pVal.FormMode;
                                    /* if (Validation() == false)
                                     {
                                         BubbleEvent = false;
                                         return false;

                                     }*/
                                    //for (int i = 1; i <= oMatrix.RowCount; i++)
                                    //{
                                    //    if (string.IsNullOrEmpty(oMatrix.Columns.Item("V_4").Cells.Item(i).Specific.value))
                                    //    {
                                    //        oMatrix.DeleteRow(i);

                                    //    }
                                    //}

                                }
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                        else
                        {
                           
                        }

                        break;

                    #endregion

                    #region LOST_FOCUS
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:

                        if (pVal.BeforeAction)
                        {
                            try
                            {

                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                        else
                        {
                        }

                        break;

                    #endregion

                    #region COMBO_SELECT
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

                        if (pVal.BeforeAction == true )
                        {
                            try
                            {

                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                        else if (pVal.BeforeAction == false)
                        {
                            if(pVal.ItemUID == "cMaingrou" || pVal.ItemUID == "Item_15")
                            {
                                fillgrid();

                            }
                           


                        }

                        break;

                    #endregion

                    #region CFL
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                        if (pVal.BeforeAction == true)
                        {
                            
                        }

                        break;

                    #endregion

                    #region CLICK
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        try
                        {
                            if (pVal.BeforeAction)
                            {
                                if (pVal.ItemUID == "cMaingrou")
                                {
                                    SecondaryScheme.FillComboFromQuery();
                                }
                                else if (pVal.ItemUID == "Item_15")
                                {
                                    SecondaryScheme.FillYearCombo();
                                }
                            }
                            else
                            {
                            }
                        }
                        catch (Exception)
                        {

                            throw;
                        }

                        break;

                        #endregion
                }
                return BubbleEvent;
            }
            catch (Exception Ex)
            {
                return false;
            }
        }
        #endregion

        #region fillgrid

        public static void fillgrid()
        {
            SAPbouiCOM.Form oForm = clsMain.SBO_Application.Forms.ActiveForm;
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_2").Specific;
            //string test = oForm.Items.Item("cMaingrou").Specific.Value.ToString();
            //string test1 = oForm.Items.Item("Item_15").Specific.Value.ToString();

            oGrid.DataTable.Rows.Clear();

            string query = "Call SP_SecondaryScheme('" + oForm.Items.Item("cMaingrou").Specific.Value.ToString() + "','" + oForm.Items.Item("Item_15").Specific.Value.ToString() + "')";
            oGrid.DataTable.ExecuteQuery(query);

            SAPbouiCOM.EditTextColumn oColumns = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Customer Code");
            oColumns.LinkedObjectType = "2";

        }

    }
    #endregion
}

