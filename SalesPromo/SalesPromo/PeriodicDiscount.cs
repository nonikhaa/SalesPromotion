using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SalesPromo
{
    public class PeriodicDiscount
    {
        private SAPbouiCOM.Application oSBOApplication;
        private SAPbobsCOM.Company oSBOCompany;

        public PeriodicDiscount(SAPbouiCOM.Application oSBOApplication, SAPbobsCOM.Company oSBOCompany)
        {
            this.oSBOApplication = oSBOApplication;
            this.oSBOCompany = oSBOCompany;
        }

        #region Menu Event
        /// <summary>
        /// Menu Event Periodic Discount
        /// When click menu, this event called
        /// </summary>
        public void MenuEvent_PrdDisc(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (pVal.BeforeAction == false)
            {
                Form oForm = null;

                try
                {
                    oForm = Utils.createForm(ref oSBOApplication, "PrdDisc");
                    oForm.Visible = true;
                    Template_Add_PrdDisc(ref oForm);
                }
                catch (Exception ex)
                {
                    bubbleEvent = false;
                    oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                finally
                {
                    if (oForm != null)
                    {
                        if (bubbleEvent)
                        {
                            oForm.Freeze(false);
                            oForm.VisibleEx = true;
                        }
                        else
                            oForm.Close();
                    }
                    Utils.releaseObject(oForm);
                }
            }
        }

        /// <summary>
        /// Template Periodic Discount When Load
        /// </summary>
        public void Template_Add_PrdDisc(ref Form oForm)
        {
            try
            {
                DBDataSource dtSource = null;
                dtSource = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_H");

                oForm.Freeze(true);
                Template_Clear_PrdDisc(ref oForm);
                oForm.DataBrowser.BrowseBy = "tCode";

                oForm.Items.Item("tab3").Click();
                oForm.Items.Item("tab2").Click();
                oForm.Items.Item("tab1").Click();

                // Generate Code
                string runCode = GenerateCode();
                dtSource.SetValue("Code", 0, runCode);
                dtSource.SetValue("U_SOL_CUSTTYPE", 0, "All Customer");
                oForm.Items.Item("mt_1").Enabled = false;

                oForm.Items.Item("tArea").Click();

                Utils.releaseObject(dtSource);
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Saat button next, previous, last record, first record di click
        /// </summary>
        public void NextPrev_PrdDisc(ref Form oForm)
        {
            try
            {
                DBDataSource dtSource = null;
                dtSource = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_H");

                if(dtSource.GetValue("U_SOL_CUSTTYPE", 0) == "Per Customer")
                    oForm.Items.Item("mt_1").Enabled = true;
                else
                    oForm.Items.Item("mt_1").Enabled = false;

                Utils.releaseObject(dtSource);
            }
            catch(Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Clear Data and Load Data Source
        /// </summary>
        private void Template_Clear_PrdDisc(ref Form oForm)
        {
            DBDataSource prdDisc_H = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_H");
            DBDataSource prdDisc_D1 = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_D1");
            DBDataSource prdDisc_D2 = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_D2");
            DBDataSource prdDisc_D3 = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_D3");

            prdDisc_H.SetValue("Code", 0, "");
            prdDisc_H.SetValue("U_SOL_AREA", 0, "");
            prdDisc_H.SetValue("U_SOL_CSTGRPCODE", 0, "");
            prdDisc_H.SetValue("U_SOL_CSTGRPNAME", 0, "");
            prdDisc_H.SetValue("U_SOL_DISCTYPE", 0, "");
            prdDisc_H.SetValue("U_SOL_STARTDATE", 0, "");
            prdDisc_H.SetValue("U_SOL_ENDDATE", 0, "");

            prdDisc_D1.Clear(); prdDisc_D2.Clear(); prdDisc_D3.Clear();

            prdDisc_D1.InsertRecord(prdDisc_D1.Size);
            prdDisc_D2.InsertRecord(prdDisc_D2.Size);
            prdDisc_D3.InsertRecord(prdDisc_D3.Size);

            prdDisc_D1.Offset = prdDisc_D1.Size - 1;
            prdDisc_D1.SetValue("LineId", prdDisc_D1.Size - 1, prdDisc_D1.Size.ToString());
            prdDisc_D2.Offset = prdDisc_D2.Size - 1;
            prdDisc_D2.SetValue("LineId", prdDisc_D2.Size - 1, prdDisc_D2.Size.ToString());
            prdDisc_D3.Offset = prdDisc_D3.Size - 1;
            prdDisc_D3.SetValue("LineId", prdDisc_D3.Size - 1, prdDisc_D3.Size.ToString());

            oForm.Items.Item("mt_1").Specific.LoadFromDataSource();
            oForm.Items.Item("mt_2").Specific.LoadFromDataSource();
            oForm.Items.Item("mt_3").Specific.LoadFromDataSource();

            Utils.releaseObject(prdDisc_H);
            Utils.releaseObject(prdDisc_D1);
            Utils.releaseObject(prdDisc_D2);
            Utils.releaseObject(prdDisc_D3);
        }

        /// <summary>
        /// Add row in Periodic Discount
        /// </summary>
        public void MenuEvent_PrdDiscAdd(ref MenuEvent pVal, ref bool bubbleEvent)
        {
            if (pVal.BeforeAction == true)
            {
                Form oForm = oSBOApplication.Forms.ActiveForm;
                if (oForm.TypeEx == "PRDDISC")
                {
                    try
                    {
                        oForm.Freeze(true);

                        string mtxName = string.Empty;
                        string dataSource = string.Empty;

                        switch (oForm.ActiveItem)
                        {
                            case "mt_1": mtxName = "mt_1"; dataSource = "@SOL_PRDDISC_D1"; break;
                            case "mt_2": mtxName = "mt_2"; dataSource = "@SOL_PRDDISC_D2"; break;
                            case "mt_3": mtxName = "mt_3"; dataSource = "@SOL_PRDDISC_D3"; break;
                        }

                        Matrix oMtx = oForm.Items.Item(mtxName).Specific;
                        oMtx.FlushToDataSource();

                        DBDataSource dtSource = oForm.DataSources.DBDataSources.Item(dataSource);
                        dtSource.InsertRecord(GeneralVariables.iDelRow);

                        if (mtxName == "mt_1")
                        {
                            dtSource.SetValue("U_SOL_CARDCODE", GeneralVariables.iDelRow, "");
                            dtSource.SetValue("U_SOL_CARDNAME", GeneralVariables.iDelRow, "");

                        }
                        else if (mtxName == "mt_2")
                        {
                            dtSource.SetValue("U_SOL_ITEMCODE", GeneralVariables.iDelRow, "");
                            dtSource.SetValue("U_SOL_ITEMNAME", GeneralVariables.iDelRow, "");
                            dtSource.SetValue("U_SOL_MINQTY", GeneralVariables.iDelRow, "0");
                            dtSource.SetValue("U_SOL_ITEMDISC", GeneralVariables.iDelRow, "0");
                        }
                        else
                        {
                            dtSource.SetValue("U_SOL_ITEMCODE", GeneralVariables.iDelRow, "");
                            dtSource.SetValue("U_SOL_ITEMNAME", GeneralVariables.iDelRow, "");
                            dtSource.SetValue("U_SOL_MINQTY", GeneralVariables.iDelRow, "0");
                            dtSource.SetValue("U_SOL_ITMCD_FREE", GeneralVariables.iDelRow, "");
                            dtSource.SetValue("U_SOL_ITMNM_FREE", GeneralVariables.iDelRow, "");
                            dtSource.SetValue("U_SOL_QTY_FREE", GeneralVariables.iDelRow, "0");
                            dtSource.SetValue("U_SOL_KELIPATAN", GeneralVariables.iDelRow, "N");
                        }
                        oMtx.LoadFromDataSource();

                        for (int i = 1; i <= oMtx.RowCount; i++)
                        {
                            oMtx.Columns.Item("#").Cells.Item(i).Specific.value = i;
                        }
                        oMtx.FlushToDataSource();

                        if (mtxName == "mt_3")
                            oMtx.Columns.Item(2).Cells.Item(GeneralVariables.iDelRow + 1).Click();
                        else
                            oMtx.Columns.Item(1).Cells.Item(GeneralVariables.iDelRow + 1).Click();

                        Utils.releaseObject(dtSource);
                        Utils.releaseObject(oMtx);
                    }
                    catch (Exception ex)
                    {
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        oForm.Freeze(false);
                        Utils.releaseObject(oForm);
                    }
                }
            }
        }

        /// <summary>
        /// Delete row in Periodic Discount
        /// </summary>
        public void MenuEvent_PrdDiscDel(ref MenuEvent pVal, ref bool bubbleEvent)
        {
            if (pVal.BeforeAction == true)
            {
                Form oForm = oSBOApplication.Forms.ActiveForm;
                if (oForm.TypeEx == "PRDDISC")
                {
                    try
                    {
                        oForm.Freeze(true);

                        string mtxName = string.Empty;
                        string dataSource = string.Empty;

                        switch (oForm.ActiveItem)
                        {
                            case "mt_1": mtxName = "mt_1"; dataSource = "@SOL_PRDDISC_D1"; break;
                            case "mt_2": mtxName = "mt_2"; dataSource = "@SOL_PRDDISC_D2"; break;
                            case "mt_3": mtxName = "mt_3"; dataSource = "@SOL_PRDDISC_D3"; break;
                        }

                        Matrix oMtx = oForm.Items.Item(mtxName).Specific;
                        oMtx.FlushToDataSource();
                        try
                        {
                            oMtx.DeleteRow(GeneralVariables.iDelRow);
                        }
                        catch (Exception ex) { }

                        for (int i = 1; i <= oMtx.RowCount; i++)
                        {
                            oMtx.Columns.Item("#").Cells.Item(i).Specific.value = i;
                        }
                        RefreshMatrix(oForm, mtxName, dataSource);

                        Utils.releaseObject(oMtx);
                    }
                    catch (Exception ex)
                    {
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        oForm.Freeze(false);
                        Utils.releaseObject(oForm);
                    }
                }
            }
        }

        /// <summary>
        /// Refresh Matrix
        /// </summary>
        private void RefreshMatrix(Form oForm, string mtxName, string dataSource)
        {
            DBDataSource dtSource = oForm.DataSources.DBDataSources.Item(dataSource);
            Matrix oMtx = oForm.Items.Item(mtxName).Specific;
            dtSource.Clear();
            oMtx.FlushToDataSource();
            Utils.releaseObject(dtSource);
            Utils.releaseObject(oMtx);
        }

        /// <summary>
        /// Generate Code
        /// </summary>
        private string GenerateCode()
        {
            string runNumber = string.Empty;
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                    oRec.DoQuery("CALL SOL_SP_ADDON_PD_CODE");
                else
                    oRec.DoQuery("EXEC SOL_SP_ADDON_PD_CODE");

                if (oRec.RecordCount > 0)
                {
                    runNumber = oRec.Fields.Item("RunNumber").Value;
                }
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
            finally
            {
                Utils.releaseObject(oRec);
            }

            return runNumber;
        }

        #endregion

        #region Right Click Event
        /// <summary>
        /// Create menu when Rihgt Click Event
        /// menu Add Row and Delete Row
        /// </summary>
        public void RightClickEvent_PrdDisc(ref ContextMenuInfo eventInfo, ref bool bubbleEvent)
        {
            Form oForm = oSBOApplication.Forms.ActiveForm;

            if (eventInfo.BeforeAction == true && (eventInfo.ItemUID == "mt_1" || eventInfo.ItemUID == "mt_2" || eventInfo.ItemUID == "mt_3"))
            {
                MenuItem oMenuItem = null;
                Menus oMenus = null;
                MenuCreationParams oCreateionPackage = null;

                try
                {
                    oCreateionPackage = oSBOApplication.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

                    oCreateionPackage.Type = BoMenuType.mt_STRING;
                    oCreateionPackage.UniqueID = "PrdDiscAdd";
                    oCreateionPackage.Position = 1;
                    oCreateionPackage.String = "Add Row";
                    oCreateionPackage.Enabled = true;

                    oMenuItem = oSBOApplication.Menus.Item("1280");
                    oMenus = oMenuItem.SubMenus;
                    oMenus.AddEx(oCreateionPackage);

                    oCreateionPackage = oSBOApplication.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

                    oCreateionPackage.Type = BoMenuType.mt_STRING;
                    oCreateionPackage.UniqueID = "PrdDiscDel";
                    oCreateionPackage.Position = 2;
                    oCreateionPackage.String = "Delete Row";
                    oCreateionPackage.Enabled = true;

                    oMenuItem = oSBOApplication.Menus.Item("1280");
                    oMenus = oMenuItem.SubMenus;
                    oMenus.AddEx(oCreateionPackage);
                }
                catch (Exception ex)
                {
                    oSBOApplication.MessageBox(ex.Message);
                }
                finally
                {
                    Utils.releaseObject(oMenuItem);
                    Utils.releaseObject(oMenus);
                    Utils.releaseObject(oCreateionPackage);
                }
                GeneralVariables.iDelRow = eventInfo.Row;
            }
            else
            {
                oSBOApplication.Menus.RemoveEx("PrdDiscAdd");
                oSBOApplication.Menus.RemoveEx("PrdDiscDel");
            }

            Utils.releaseObject(oForm);
        }
        #endregion;

        #region Item Event
        /// <summary>
        /// Item Event Periodic Discount
        /// </summary>
        public void ItemEvent_PrdDisc(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            switch (pVal.EventType)
            {
                case BoEventTypes.et_FORM_LOAD:; break;
                case BoEventTypes.et_CHOOSE_FROM_LIST: CFL_PrdDisc(formUID, ref pVal, ref bubbleEvent); break;
                case BoEventTypes.et_VALIDATE: Validate_PrdDisc(formUID, ref pVal, ref bubbleEvent); break;
                case BoEventTypes.et_COMBO_SELECT: Combo_PrdDisc(formUID, ref pVal, ref bubbleEvent); break;
            }
        }
        /// <summary>
        /// Choose From List
        /// </summary>
        private void CFL_PrdDisc(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            IChooseFromListEvent oCFLEvent = null;
            try
            {
                oCFLEvent = (IChooseFromListEvent)pVal;

                switch (pVal.ItemUID)
                {
                    case "tGrpCd": CFL_PrdDisc_CustGrp(formUID, ref pVal, ref bubbleEvent, ref oCFLEvent); break;
                }
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
            finally
            {
                Utils.releaseObject(oCFLEvent);
            }
        }

        /// <summary>
        /// CFL Customer group
        /// </summary>
        private void CFL_PrdDisc_CustGrp(string formUID, ref ItemEvent pVal,
                                        ref bool bubbleEvent, ref IChooseFromListEvent oCFLEvent)
        {
            if (bubbleEvent)
            {
                Form oForm = null;
                Conditions oCons = null;
                ICondition oCon = null;
                SAPbouiCOM.ChooseFromList oCFL = null;
                DataTable oDataTable = null;
                BusinessPartnerGroups oBpGrp = null;
                DBDataSource oDBSource_H = null;

                try
                {
                    oForm = oSBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                    oBpGrp = oSBOCompany.GetBusinessObject(BoObjectTypes.oBusinessPartnerGroups);

                    if (oCFLEvent.BeforeAction == true)
                    {
                        if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            oForm.Freeze(true);
                            oCFL = oForm.ChooseFromLists.Item("CFL_CSTGRP");

                            oCFL.SetConditions(null);
                            oCons = oCFL.GetConditions();

                            oCon = oCons.Add();
                            oCon.Alias = "GroupType";
                            oCon.Operation = BoConditionOperation.co_EQUAL;
                            oCon.CondVal = "C";

                            oCFL.SetConditions(oCons);
                        }
                    }
                    else if (oCFLEvent.BeforeAction == false && oCFLEvent.ActionSuccess == true && oCFLEvent.SelectedObjects != null)
                    {
                        oForm.Freeze(true);
                        oDataTable = oCFLEvent.SelectedObjects;

                        if (oBpGrp.GetByKey(oDataTable.GetValue("GroupCode", 0)))
                        {
                            oDBSource_H = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_H");
                            oDBSource_H.SetValue("U_SOL_CSTGRPCODE", 0, oBpGrp.Code.ToString());
                            oDBSource_H.SetValue("U_SOL_CSTGRPNAME", 0, oBpGrp.Name);

                            if (oForm.Mode != BoFormMode.fm_ADD_MODE)
                                oForm.Mode = BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }
                catch (Exception ex)
                {
                    bubbleEvent = false;
                    oSBOApplication.MessageBox(ex.Message + " : " + ex.StackTrace);
                }
                finally
                {
                    if (oForm != null) oForm.Freeze(false);

                    Utils.releaseObject(oForm);
                    Utils.releaseObject(oCons);
                    Utils.releaseObject(oCon);
                    Utils.releaseObject(oCFL);
                    Utils.releaseObject(oDataTable);
                    Utils.releaseObject(oBpGrp);
                    Utils.releaseObject(oDBSource_H);
                }
            }
        }

       

        #region Validate
        /// <summary>
        /// Validate Event
        /// </summary>
        private void Validate_PrdDisc(string formUID, ref ItemEvent pVal,ref bool bubbleEvent)
        {
            switch (pVal.ColUID)
            {
                case "cCsCd": validate_PrdDisc_CustCode(formUID, ref pVal, ref bubbleEvent); break;
                case "cItmCd": validate_PrdDisc_ItemCodePrcnt(formUID, ref pVal, ref bubbleEvent); break;
                case "cItmNm": validate_PrdDisc_ItemNamePrcnt(formUID, ref pVal, ref bubbleEvent); break;
                case "cItmC": validate_PrdDisc_ItemCodeB1G1(formUID, ref pVal, ref bubbleEvent); break;
                case "cItmN": validate_PrdDisc_ItemNameB1G1(formUID, ref pVal, ref bubbleEvent); break;
                case "cItmCB": validate_PrdDisc_ItemCodeBns(formUID, ref pVal, ref bubbleEvent); break;
                case "cItmNB": validate_PrdDisc_ItemNameBns(formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        /// <summary>
        /// Validate - Customer Code in tab Customer
        /// </summary>
        private void validate_PrdDisc_CustCode(string formUID, ref ItemEvent pVal,ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    Matrix oMtx = oForm.Items.Item("mt_1").Specific;
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_D1");
                    Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    String cardCode = oMtx.Columns.Item("cCsCd").Cells.Item(pVal.Row).Specific.Value;

                    try
                    {
                        oForm.Freeze(true);
                        oMtx.FlushToDataSource();

                        if (pVal.Row == oMtx.RowCount)
                        {
                            dtSource.InsertRecord(dtSource.Size);
                            dtSource.SetValue("LineId", dtSource.Size - 1, dtSource.Size.ToString());
                        }

                        if (cardCode != "ALL")
                        {
                            string nama = string.Empty;
                            oRec.DoQuery("SELECT \"CardName\" FROM OCRD WHERE \"CardCode\" = '"
                                         + cardCode + "'");

                            if (oRec.RecordCount > 0)
                            {
                                nama = oRec.Fields.Item(0).Value;
                            }

                            dtSource.SetValue("U_SOL_CARDNAME", pVal.Row - 1, nama);
                        }
                        else
                        {
                            dtSource.SetValue("U_SOL_CARDNAME", pVal.Row - 1, "All Customer");
                        }

                        oMtx.LoadFromDataSource();
                    }
                    catch (Exception ex)
                    {
                        bubbleEvent = false;
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        if (oForm != null) oForm.Freeze(false);

                        Utils.releaseObject(oForm);
                        Utils.releaseObject(oMtx);
                        Utils.releaseObject(dtSource);
                        Utils.releaseObject(oRec);
                    }
                }
            }
        }

        

        /// <summary>
        /// Validate - Item Code in tab Discount %
        /// </summary>
        private void validate_PrdDisc_ItemCodePrcnt(string formUID, ref ItemEvent pVal,ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    Matrix oMtx = oForm.Items.Item("mt_2").Specific;
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_D2");
                    Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    try
                    {
                        oForm.Freeze(true);
                        oMtx.FlushToDataSource();

                        if (pVal.Row == oMtx.RowCount)
                        {
                            dtSource.InsertRecord(dtSource.Size);
                            dtSource.SetValue("LineId", dtSource.Size - 1, dtSource.Size.ToString());
                        }

                        string nama = string.Empty;
                        oRec.DoQuery("SELECT \"ItemName\" FROM OITM WHERE \"ItemCode\" = '"
                                     + oMtx.Columns.Item("cItmCd").Cells.Item(pVal.Row).Specific.Value + "'");

                        if (oRec.RecordCount > 0)
                        {
                            nama = oRec.Fields.Item(0).Value;
                        }

                        dtSource.SetValue("U_SOL_ITEMNAME", pVal.Row - 1, nama);

                        oMtx.LoadFromDataSource();
                    }
                    catch (Exception ex)
                    {
                        bubbleEvent = false;
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        if (oForm != null) oForm.Freeze(false);

                        Utils.releaseObject(oForm);
                        Utils.releaseObject(oMtx);
                        Utils.releaseObject(dtSource);
                        Utils.releaseObject(oRec);
                    }
                }
            }
        }

        /// <summary>
        /// Validate - Item Name in tab Discount %
        /// </summary>
        private void validate_PrdDisc_ItemNamePrcnt(string formUID, ref ItemEvent pVal,ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    Matrix oMtx = oForm.Items.Item("mt_2").Specific;
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_D2");
                    Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    try
                    {
                        oForm.Freeze(true);
                        oMtx.FlushToDataSource();

                        if (pVal.Row == oMtx.RowCount)
                        {
                            dtSource.InsertRecord(dtSource.Size);
                            dtSource.SetValue("LineId", dtSource.Size - 1, dtSource.Size.ToString());
                        }

                        string code = string.Empty;
                        oRec.DoQuery("SELECT \"ItemCode\" FROM OITM WHERE \"ItemName\" = '"
                                     + oMtx.Columns.Item("cItmNm").Cells.Item(pVal.Row).Specific.Value + "'");

                        if (oRec.RecordCount > 0)
                        {
                            code = oRec.Fields.Item(0).Value;
                        }

                        dtSource.SetValue("U_SOL_ITEMCODE", pVal.Row - 1, code);

                        oMtx.LoadFromDataSource();
                    }
                    catch (Exception ex)
                    {
                        bubbleEvent = false;
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        if (oForm != null) oForm.Freeze(false);

                        Utils.releaseObject(oForm);
                        Utils.releaseObject(oMtx);
                        Utils.releaseObject(dtSource);
                        Utils.releaseObject(oRec);
                    }
                }
            }
        }

        /// <summary>
        /// Validate - Item Code in tab Buy 1 Get 1
        /// </summary>
        private void validate_PrdDisc_ItemCodeB1G1(string formUID, ref ItemEvent pVal,ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    Matrix oMtx = oForm.Items.Item("mt_3").Specific;
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_D3");
                    Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    try
                    {
                        oForm.Freeze(true);
                        oMtx.FlushToDataSource();

                        if (pVal.Row == oMtx.RowCount)
                        {
                            dtSource.InsertRecord(dtSource.Size);
                            dtSource.SetValue("LineId", dtSource.Size - 1, dtSource.Size.ToString());
                        }

                        string nama = string.Empty;
                        oRec.DoQuery("SELECT \"ItemName\" FROM OITM WHERE \"ItemCode\" = '"
                                     + oMtx.Columns.Item("cItmC").Cells.Item(pVal.Row).Specific.Value + "'");

                        if (oRec.RecordCount > 0)
                        {
                            nama = oRec.Fields.Item(0).Value;
                        }

                        dtSource.SetValue("U_SOL_ITEMNAME", pVal.Row - 1, nama);

                        oMtx.LoadFromDataSource();
                    }
                    catch (Exception ex)
                    {
                        bubbleEvent = false;
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        if (oForm != null) oForm.Freeze(false);

                        Utils.releaseObject(oForm);
                        Utils.releaseObject(oMtx);
                        Utils.releaseObject(dtSource);
                        Utils.releaseObject(oRec);
                    }
                }
            }
        }

        /// <summary>
        /// Validate - Item Name in tab Buy 1 Get 1
        /// </summary>
        private void validate_PrdDisc_ItemNameB1G1(string formUID, ref ItemEvent pVal,ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    Matrix oMtx = oForm.Items.Item("mt_3").Specific;
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_D3");
                    Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    try
                    {
                        oForm.Freeze(true);
                        oMtx.FlushToDataSource();

                        if (pVal.Row == oMtx.RowCount)
                        {
                            dtSource.InsertRecord(dtSource.Size);
                            dtSource.SetValue("LineId", dtSource.Size - 1, dtSource.Size.ToString());
                        }

                        string code = string.Empty;
                        oRec.DoQuery("SELECT \"ItemCode\" FROM OITM WHERE \"ItemName\" = '"
                                     + oMtx.Columns.Item("cItmN").Cells.Item(pVal.Row).Specific.Value + "'");

                        if (oRec.RecordCount > 0)
                        {
                            code = oRec.Fields.Item(0).Value;
                        }

                        dtSource.SetValue("U_SOL_ITEMCODE", pVal.Row - 1, code);

                        oMtx.LoadFromDataSource();
                    }
                    catch (Exception ex)
                    {
                        bubbleEvent = false;
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        if (oForm != null) oForm.Freeze(false);

                        Utils.releaseObject(oForm);
                        Utils.releaseObject(oMtx);
                        Utils.releaseObject(dtSource);
                        Utils.releaseObject(oRec);
                    }
                }
            }
        }

        /// <summary>
        /// Validate - Item Code Bonus in tab Buy 1 Get 1
        /// </summary>
        private void validate_PrdDisc_ItemCodeBns(string formUID, ref ItemEvent pVal,ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    Matrix oMtx = oForm.Items.Item("mt_3").Specific;
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_D3");
                    Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    try
                    {
                        oForm.Freeze(true);
                        oMtx.FlushToDataSource();

                        if (pVal.Row == oMtx.RowCount)
                        {
                            dtSource.InsertRecord(dtSource.Size);
                            dtSource.SetValue("LineId", dtSource.Size - 1, dtSource.Size.ToString());
                        }

                        string nama = string.Empty;
                        oRec.DoQuery("SELECT \"ItemName\" FROM OITM WHERE \"ItemCode\" = '"
                                     + oMtx.Columns.Item("cItmCB").Cells.Item(pVal.Row).Specific.Value + "'");

                        if (oRec.RecordCount > 0)
                        {
                            nama = oRec.Fields.Item(0).Value;
                        }

                        dtSource.SetValue("U_SOL_ITMNM_FREE", pVal.Row - 1, nama);

                        oMtx.LoadFromDataSource();
                    }
                    catch (Exception ex)
                    {
                        bubbleEvent = false;
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        if (oForm != null) oForm.Freeze(false);

                        Utils.releaseObject(oForm);
                        Utils.releaseObject(oMtx);
                        Utils.releaseObject(dtSource);
                        Utils.releaseObject(oRec);
                    }
                }
            }
        }

        /// <summary>
        /// Validate - Item Name Bonus in tab Buy 1 Get 1
        /// </summary>
        private void validate_PrdDisc_ItemNameBns(string formUID, ref ItemEvent pVal,ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    Matrix oMtx = oForm.Items.Item("mt_3").Specific;
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_PRDDISC_D3");
                    Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    try
                    {
                        oForm.Freeze(true);
                        oMtx.FlushToDataSource();

                        if (pVal.Row == oMtx.RowCount)
                        {
                            dtSource.InsertRecord(dtSource.Size);
                            dtSource.SetValue("LineId", dtSource.Size - 1, dtSource.Size.ToString());
                        }

                        string code = string.Empty;
                        oRec.DoQuery("SELECT \"ItemCode\" FROM OITM WHERE \"ItemName\" = '"
                                     + oMtx.Columns.Item("cItmNB").Cells.Item(pVal.Row).Specific.Value + "'");

                        if (oRec.RecordCount > 0)
                        {
                            code = oRec.Fields.Item(0).Value;
                        }

                        dtSource.SetValue("U_SOL_ITMCD_FREE", pVal.Row - 1, code);

                        oMtx.LoadFromDataSource();
                    }
                    catch (Exception ex)
                    {
                        bubbleEvent = false;
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        if (oForm != null) oForm.Freeze(false);

                        Utils.releaseObject(oForm);
                        Utils.releaseObject(oMtx);
                        Utils.releaseObject(dtSource);
                        Utils.releaseObject(oRec);
                    }
                }
            }
        }

        #endregion

        

        /// <summary>
        /// Combo box select
        /// </summary>
        private void Combo_PrdDisc(string formUID, ref ItemEvent pVal,ref bool bubbleEvent)
        {
            switch (pVal.ItemUID)
            {
                case "cbCustType": Combo_PrdDisc_CustType(formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        /// <summary>
        /// Combo box - Customer type in tab customer
        /// </summary>
        private void Combo_PrdDisc_CustType(string formUID, ref ItemEvent pVal,ref bool bubbleEvent)
        {

            if (pVal.BeforeAction == true && pVal.ItemChanged == true)
            {
                Form oForm = oSBOApplication.Forms.Item(formUID);
                Matrix oMtx = oForm.Items.Item("mt_1").Specific;
                if (oMtx.RowCount > 1 && oMtx.Columns.Item("cCsCd").Cells.Item(1).Specific.Value != "")
                {
                    if (oSBOApplication.MessageBox("Mengganti mode customer akan menghapus isi dari tabel Customer. Lanjutkan?", 1, "Ya", "Tidak") != 1)
                    {
                        bubbleEvent = false;
                    }
                }

                Utils.releaseObject(oForm);
                Utils.releaseObject(oMtx);
            }

            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    Matrix oMtx = oForm.Items.Item("mt_1").Specific;

                    oMtx.Clear();
                    oMtx.AddRow();
                    oMtx.Columns.Item("#").Cells.Item(1).Specific.Value = 1;
                    oMtx.FlushToDataSource();
                    oForm.Items.Item("tab1").Click();

                    string custType = oForm.Items.Item("cbCustType").Specific.Value;

                    if (custType == "Per Customer")
                        oForm.Items.Item("mt_1").Enabled = true;
                    else
                        oForm.Items.Item("mt_1").Enabled = false;

                    Utils.releaseObject(oForm);
                    Utils.releaseObject(oMtx);
                }
            }
        }

        #endregion
    }
}
