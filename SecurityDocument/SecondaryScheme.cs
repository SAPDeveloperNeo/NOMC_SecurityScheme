using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SecurityScheme
{
    class SecondaryScheme
    {
        #region Variable

        private SAPbouiCOM.Form oForm, oForm1, oFormCFL;
        private SAPbouiCOM.Button oBtn = null;
        private SAPbouiCOM.Item oItem, oItem1, oItem2, oItem3;

        private SAPbouiCOM.Grid oGrid;

        private SAPbouiCOM.Matrix oMatrix, oMat;
        private Boolean ACTION = false;
        private SAPbobsCOM.Recordset oRecordSet, oRec1, oRec, oRecordSetRP;
        private int Mode;
        private int i, DelLine;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;
        private SAPbouiCOM.ChooseFromList oCFL = null;
        public static bool CloseFlg, BtnFlag = false, DBSetupFlg = false;


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
                    oForm.Items.Item("tSCode").Enabled = true;
                    oForm.Items.Item("cMaingrou").Enabled = true;
                    oForm.Items.Item("Item_3").Enabled = true;
                    oForm.Items.Item("Item_15").Enabled = true;
                    oForm.Items.Item("Item_6").Enabled = true;
                    oForm.Items.Item("Item_7").Enabled = false;
                    oForm.Items.Item("tRmk").Click();
                    //return true;
                }
                else if (Type == "Find")
                {
                    oForm.Items.Item("tSCode").Enabled = true;
                    oForm.Items.Item("cMaingrou").Enabled = true;
                    oForm.Items.Item("Item_3").Enabled = true;
                    oForm.Items.Item("Item_15").Enabled = true;
                    oForm.Items.Item("Item_6").Enabled = true;
                    oForm.Items.Item("Item_7").Enabled = false;

                    oForm.Items.Item("tCode").Specific.value = null;
                    oForm.Items.Item("tRmk").Click();
                    //return true;
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
                                if (pVal.ItemUID == "1" && (pVal.FormMode == 3)) //pVal.FormMode == 3 ||
                                {
                                    Mode = pVal.FormMode;
                                    if (Validation() == false)
                                    {
                                        BubbleEvent = false;
                                        return false;

                                    }

                                    var customercode = oForm.Items.Item("tSCode").Specific.Value;
                                    var project = oForm.Items.Item("cMaingrou").Specific.Value;
                                    var yrs = oForm.Items.Item("Item_15").Specific.Value;

                                    string queryrmnamt = "Call SP_UPDATEARTRMAMT('" + customercode + "','" + project + "','" + yrs + "')";
                                    oRecordSet.DoQuery(queryrmnamt);
                                    return true;

                                }
                                else if (pVal.ItemUID == "1" && (pVal.FormMode == 2 || pVal.FormMode == 3))
                                {
                                    oRecordSet = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    string queryV = "Select \"U_SPoj\",\"U_SMonth\",\"U_SYear\",\"U_SAct\" From \"@OCSH\"";
                                    oRecordSet.DoQuery(queryV);
                                    if (!oRecordSet.EoF)
                                    {
                                        for (int i = 1; i <= oRecordSet.RecordCount; i++)
                                        {
                                            if (oForm.Items.Item("cMaingrou").Specific.Value.ToString() == oRecordSet.Fields.Item("U_SPoj").Value.ToString() &&
                                                oForm.Items.Item("Item_3").Specific.Value.ToString() == oRecordSet.Fields.Item("U_SMonth").Value.ToString() &&
                                                oForm.Items.Item("Item_15").Specific.Value.ToString() == oRecordSet.Fields.Item("U_SYear").Value.ToString() &&
                                                oForm.Items.Item("Item_6").Specific.Value.ToString() == oRecordSet.Fields.Item("U_SAct").Value.ToString()

                                                )
                                            {
                                                var customercode = oForm.Items.Item("tSCode").Specific.Value;
                                                var project = oForm.Items.Item("cMaingrou").Specific.Value;
                                                var yrs = oForm.Items.Item("Item_15").Specific.Value;

                                                string queryrmnamt = "Call SP_UPDATEARTRMAMT('" + customercode + "','" + project + "','" + yrs + "')";
                                                oRecordSet.DoQuery(queryrmnamt);
                                                return true;

                                                //var test = oRecordSet.Fields.Item("U_SInactive").Value.ToString();

                                                //if (oRecordSet.Fields.Item("U_SInactive").Value.ToString() == "Y")
                                                //{
                                                //    return true;
                                                //}

                                                //else if (oRecordSet.Fields.Item("U_SInactive").Value.ToString() == "N")
                                                //{
                                                //    SAPbouiCOM.CheckBox checkBox1 = (SAPbouiCOM.CheckBox)oForm.Items.Item("SInactive").Specific;
                                                //    checkBox1.Checked = true;
                                                //    clsMain.SBO_Application.StatusBar.SetSystemMessage("You Can not update Data Already Exits!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                                //    //return false;
                                                //    BubbleEvent = false;


                                                //}


                                            }

                                            oRecordSet.MoveNext();

                                        }

                                    }

                                  

                                }
                               
                            }

                            catch (Exception ex)
                            {
                                throw;
                            }
                        }
                        
                        //}

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
                            try
                            {

                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }

                        break;

                    #endregion

                    #region COMBO_SELECT
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

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

                    #region CFL
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "tSCode")
                            {
                                CFLCondition("CFL_SCSH");
                            }

                        }

                        else if (pVal.BeforeAction == false)
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                            oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                            string sCFL_ID = null;
                            sCFL_ID = oCFLEvento.ChooseFromListUID;
                            string val1 = null;
                            SAPbouiCOM.ChooseFromList oCFL = null;
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                            SAPbouiCOM.DataTable oDataTable = null;
                            oDataTable = oCFLEvento.SelectedObjects;
                            if (pVal.ItemUID == "tSCode")
                            {
                                try
                                {

                                    oForm.Items.Item("tSCode").Specific.Value = oDataTable.GetValue("CardCode", 0).ToString();
                                    oForm.Items.Item("tName").Specific.Value = oDataTable.GetValue("CardName", 0).ToString();
                                }
                                catch { }
                                try
                                {

                                    
                                }
                                catch { }
                            }

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

                                }
                                else if (pVal.ItemUID == "Item_15")
                                {

                                }
                                else if (pVal.ItemUID == "Item_7" && oForm.Items.Item("Item_7").Enabled == true)
                                {
                                    OpenUserDefinedForm();
                                }
                                else if (pVal.ItemUID == "Item_3")
                                {

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

                if (pVal.BeforeAction == true)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {

                        if (CloseFlg == true)
                        {
                            BubbleEvent = false;
                            CloseFlg = false;
                            int Msg = clsMain.SBO_Application.MessageBox("Do you want to Save Changes ?", 1, "Yes", "No", "Cancel");
                            if (Msg == 1 || Msg == 2)
                            {
                                clsMain.SBO_Application.Forms.Item(FormId).Items.Item("2").Click();
                            }

                        }


                    }
  
                }
                return BubbleEvent;
            }
            catch (Exception Ex)
            {
                // oForm.Freeze(false);
                return false;
            }
        }
        #endregion

        #region FormDataEvent
        public void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {

            try
            {
                oForm = clsMain.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                string psFormId = BusinessObjectInfo.FormUID;
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        if (BusinessObjectInfo.BeforeAction == true)
                        {


                        }
                        else
                        {
                            oForm.Items.Item("tSCode").Enabled = false;
                            oForm.Items.Item("cMaingrou").Enabled = false;
                            oForm.Items.Item("Item_6").Enabled = false;
                            oForm.Items.Item("Item_15").Enabled = false;
                            oForm.Items.Item("Item_4").Enabled = false;
                            oForm.Items.Item("Item_3").Enabled = false;
                            oForm.Items.Item("Item_7").Enabled = true;

                        }

                        break;
                }
            }
            catch { }
        }
        #endregion

        #region Validation
        public bool Validation()
        {
            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("tSCode").Specific.Value)){
                    clsMain.SBO_Application.StatusBar.SetSystemMessage("Card Code Could Not Empty!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("tSCode").Click();
                    return false;

                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("cMaingrou").Specific.Value))
                {
                    clsMain.SBO_Application.StatusBar.SetSystemMessage("Project Could Not Empty!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("cMaingrou").Click();
                    return false;

                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("Item_3").Specific.Value))
                {
                    clsMain.SBO_Application.StatusBar.SetSystemMessage("Month Could Not Empty!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("Item_3").Click();
                    return false;

                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("Item_15").Specific.Value))
                {
                    clsMain.SBO_Application.StatusBar.SetSystemMessage("Year Could Not Empty!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("Item_15").Click();
                    return false;

                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("Item_6").Specific.Value))
                {
                    clsMain.SBO_Application.StatusBar.SetSystemMessage("Activity Could Not Empty!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("Item_6").Click();
                    return false;

                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("Item_14").Specific.Value))
                {
                    clsMain.SBO_Application.StatusBar.SetSystemMessage(" Claim Amount Could Not Empty!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("Item_14").Click();
                    return false;

                }
                oRecordSet = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string queryV = "Select Distinct \"U_SCode\", \"U_SPoj\",\"U_SMonth\",\"U_SYear\",\"U_SAct\" From \"@OCSH\"";
                oRecordSet.DoQuery(queryV);
                if (!oRecordSet.EoF)
                {
                    for (int i = 1; i <= oRecordSet.RecordCount; i++)
                    {
                        if(oForm.Items.Item("tSCode").Specific.Value.ToString() == oRecordSet.Fields.Item("U_SCode").Value.ToString() &&
                            oForm.Items.Item("cMaingrou").Specific.Value.ToString() == oRecordSet.Fields.Item("U_SPoj").Value.ToString() &&
                            oForm.Items.Item("Item_3").Specific.Value.ToString() == oRecordSet.Fields.Item("U_SMonth").Value.ToString() &&
                            oForm.Items.Item("Item_15").Specific.Value.ToString() == oRecordSet.Fields.Item("U_SYear").Value.ToString() 
                            )
                        {
                            clsMain.SBO_Application.StatusBar.SetSystemMessage("This Customer have already have Activity in this Month!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                            return false;
                        }


                        //else if (oForm.Items.Item("cMaingrou").Specific.Value.ToString() == oRecordSet.Fields.Item("U_SPoj").Value.ToString() &&
                        //    oForm.Items.Item("Item_3").Specific.Value.ToString() == oRecordSet.Fields.Item("U_SMonth").Value.ToString() &&
                        //    oForm.Items.Item("Item_15").Specific.Value.ToString() == oRecordSet.Fields.Item("U_SYear").Value.ToString() &&
                        //    oForm.Items.Item("Item_6").Specific.Value.ToString() == oRecordSet.Fields.Item("U_SAct").Value.ToString()

                        //    ) 
                        //{
                        //        clsMain.SBO_Application.StatusBar.SetSystemMessage("Data Already Exits!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        return false;

                        //}

                        oRecordSet.MoveNext();

                    }
                  

                }


            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Validation : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
            return true;
        }
        #endregion


        #region CFL
        public object CFLCondition(string CFL)
        {
            try
            {
                oCFL = oForm.ChooseFromLists.Item(CFL);
                oConds = clsMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                oRecordSet = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (CFL == "CFL_SCSH")
                {
                    oRecordSet.DoQuery("SELECT \"CardCode\" FROM \"OCRD\" WHERE \"CardType\"='C'AND \"validFor\" = 'Y'");
                    if (oRecordSet.RecordCount > 0)
                    {
                        for (int i = 1; i <= oRecordSet.RecordCount; i++)
                        {
                            oCond = oConds.Add();
                            oCond.Alias = "CardCode";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCond.CondVal = oRecordSet.Fields.Item(0).Value.ToString();
                            if (i != oRecordSet.RecordCount)
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                            oRecordSet.MoveNext();
                        }
                        oCFL.SetConditions(oConds);
                        return true;
                    }
                    else
                    {
                        clsMain.SBO_Application.SetStatusBarMessage("No Record found", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        oCond = oConds.Add();
                        oCond.Alias = "CardCode";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCond.CondVal = null;
                        oCFL.SetConditions(oConds);
                        return true;
                    }

                }
                return true;

            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("CFL : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
        }
        #endregion

        #region fillcomboMonth
        public static void fillcomboMonth()
        {
            try
            {
                SAPbouiCOM.Form oForm;
                SAPbobsCOM.Recordset oRec;
                SAPbouiCOM.ComboBox oCombo;
                oForm = clsMain.SBO_Application.Forms.ActiveForm;
                oCombo = null;
                oCombo = oForm.Items.Item("Item_3").Specific;
                if (oCombo.ValidValues.Count == 0)
                {
                    oCombo.ValidValues.Add("1", "January");
                    oCombo.ValidValues.Add("2", "February");
                    oCombo.ValidValues.Add("3", "March");
                    oCombo.ValidValues.Add("4", "April");
                    oCombo.ValidValues.Add("5", "May");
                    oCombo.ValidValues.Add("6", "June");
                    oCombo.ValidValues.Add("7", "July");
                    oCombo.ValidValues.Add("8", "August");
                    oCombo.ValidValues.Add("9", "September");
                    oCombo.ValidValues.Add("10", "October");
                    oCombo.ValidValues.Add("11", "November");
                    oCombo.ValidValues.Add("12", "December");

                }
                oForm.Items.Item("Item_3").DisplayDesc = true;
                oCombo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
                // oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

            }
            catch (Exception ex)
            {

            }
        }
        #endregion

        #region FillComboFromQuery
        public static void FillComboFromQuery()
        {
            try
            {
                SAPbouiCOM.Form oForm = clsMain.SBO_Application.Forms.ActiveForm;
                SAPbouiCOM.ComboBox oCombo = oForm.Items.Item("cMaingrou").Specific;

                if (oCombo.ValidValues.Count == 0)
                {
                    SAPbobsCOM.Recordset oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string query = "SELECT DISTINCT \"PrjCode\", \"PrjName\" FROM OPRJ";
                    oRec.DoQuery(query);

                    if (!oRec.EoF)
                    {
                        while (!oRec.EoF)
                        {
                            string prjCode = oRec.Fields.Item("PrjCode").Value.ToString();
                            string prjName = oRec.Fields.Item("PrjName").Value.ToString();

                            oCombo.ValidValues.Add(prjCode, prjName);

                            oRec.MoveNext();
                        }
                    }
                }

                oForm.Items.Item("cMaingrou").DisplayDesc = true;
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                //oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            }
            catch (Exception ex)
            {
                // Handle the exception appropriately, e.g., display an error message
                Console.WriteLine("Exception: " + ex.Message);
            }
        }
        #endregion

        #region FillYearCombo
        public static void FillYearCombo()
        {
            try
            {
                SAPbouiCOM.Form oForm = clsMain.SBO_Application.Forms.ActiveForm;
                SAPbouiCOM.ComboBox oCombo = oForm.Items.Item("Item_15").Specific;

                if (oCombo.ValidValues.Count == 0)
                {
                    SAPbobsCOM.Recordset oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string query = "SELECT DISTINCT YEAR(\"F_RefDate\")AS F_RefYear, YEAR(\"F_RefDate\") AS F_RefYear1 FROM OFPR ORDER BY F_RefYear DESC";
                    oRec.DoQuery(query);

                    if (!oRec.EoF)
                    {
                        while (!oRec.EoF)
                        {
                            string Year = oRec.Fields.Item("F_RefYear").Value.ToString();
                            string Year1 = oRec.Fields.Item("F_RefYear1").Value.ToString();

                            oCombo.ValidValues.Add(Year, Year1);

                            oRec.MoveNext();
                        }
                    }
                }

                oForm.Items.Item("Item_15").DisplayDesc = true;
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                //oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            }
            catch (Exception ex)
            {
                // Handle the exception appropriately, e.g., display an error message
                Console.WriteLine("Exception: " + ex.Message);
            }
        }
        #endregion

        #region OpenUserDefinedForm

        public void OpenUserDefinedForm()
        {
           
            try
            {
                SAPbouiCOM.Form oForm = clsMain.SBO_Application.Forms.ActiveForm;
                SAPbobsCOM.Recordset oRecordSetRP = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string test = oForm.Items.Item("tCode").Specific.value;
                string queryCardfillgirid = "Select Distinct B.\"CardCode\",B.\"Balance\" From \"@OCSH\" A inner join OCRD B on A.\"U_SCode\" = B.\"CardCode\" Where A.\"U_SCode\" = '"+ oForm.Items.Item("tCode").Specific.value + "' and B.\"CardType\" = 'C'";

                oRecordSetRP.DoQuery(queryCardfillgirid);

                string fileName = "frmSecondarySchemeReport";

                clsMain.LoadFromXML(fileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }
        }
        #endregion
    }
}

