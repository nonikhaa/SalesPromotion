using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SalesPromo
{
    public class FixDiscount
    {
        #region Menu Event
        /// <summary>
        /// Menu Event Fix Discount
        /// When click menu, this event called
        /// </summary>
        public void MenuEvent_FixDisc(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication
                                       , ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (pVal.BeforeAction == false)
            {
                Form oForm = null;

                try
                {
                    oForm = Utils.createForm(ref oSBOApplication, "FixDisc");
                    oForm.Visible = true;
                    Template_Add_FixDisc(ref oSBOCompany, ref oSBOApplication, ref oForm);
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
        /// Template Fix Discount When Load
        /// </summary>
        public void Template_Add_FixDisc(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication, ref Form oForm)
        {
            try
            {
                DBDataSource dtSource = null;
                dtSource = oForm.DataSources.DBDataSources.Item("@SOL_FIXDISC_H");

                oForm.Freeze(true);
                Template_Clear_FixDisc(ref oForm);
                oForm.DataBrowser.BrowseBy = "tCode";

                oForm.Items.Item("tab2").Click();
                oForm.Items.Item("tab1").Click();
                
                oForm.Items.Item("tCustCd").Click();

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
        /// Clear Data and Load Data Source
        /// </summary>
        private void Template_Clear_FixDisc(ref Form oForm)
        {
            DBDataSource fixDisc_H = oForm.DataSources.DBDataSources.Item("@SOL_FIXDISC_H");
            DBDataSource fixDisc_D1 = oForm.DataSources.DBDataSources.Item("@SOL_FIXDISC_D1");
            DBDataSource fixDisc_D2 = oForm.DataSources.DBDataSources.Item("@SOL_FIXDISC_D2");

            fixDisc_H.SetValue("Code", 0, "");
            fixDisc_H.SetValue("U_SOL_AREA", 0, "");
            fixDisc_H.SetValue("U_SOL_CARDCODE", 0, "");
            fixDisc_H.SetValue("U_SOL_CARDNAME", 0, "");

            fixDisc_D1.Clear(); fixDisc_D1.Clear();

            fixDisc_D1.InsertRecord(fixDisc_D1.Size);
            fixDisc_D2.InsertRecord(fixDisc_D2.Size);

            fixDisc_D1.Offset = fixDisc_D1.Size - 1;
            fixDisc_D1.SetValue("LineId", fixDisc_D1.Size - 1, fixDisc_D1.Size.ToString());
            fixDisc_D2.Offset = fixDisc_D2.Size - 1;
            fixDisc_D2.SetValue("LineId", fixDisc_D2.Size - 1, fixDisc_D2.Size.ToString());

            oForm.Items.Item("mt_1").Specific.LoadFromDataSource();
            oForm.Items.Item("mt_2").Specific.LoadFromDataSource();

            Utils.releaseObject(fixDisc_H);
            Utils.releaseObject(fixDisc_D1);
            Utils.releaseObject(fixDisc_D2);
        }

        /// <summary>
        /// Add row in Fix Discount
        /// </summary>
        public void MenuEvent_FixDiscAdd(ref Application oSBOApplication, ref MenuEvent pVal, ref bool bubbleEvent)
        {
            if (pVal.BeforeAction == true)
            {
                Form oForm = oSBOApplication.Forms.ActiveForm;
                if (oForm.TypeEx == "FIXDISC")
                {
                    try
                    {
                        oForm.Freeze(true);

                        string mtxName = string.Empty;
                        string dataSource = string.Empty;

                        switch (oForm.ActiveItem)
                        {
                            case "mt_1": mtxName = "mt_1"; dataSource = "@SOL_FIXDISC_D1"; break;
                            case "mt_2": mtxName = "mt_2"; dataSource = "@SOL_FIXDISC_D2"; break;
                        }

                        Matrix oMtx = oForm.Items.Item(mtxName).Specific;
                        oMtx.FlushToDataSource();

                        DBDataSource dtSource = oForm.DataSources.DBDataSources.Item(dataSource);
                        dtSource.InsertRecord(GeneralVariables.iDelRow);

                        if (mtxName == "mt_1")
                        {
                            dtSource.SetValue("U_SOL_BRANDCODE", GeneralVariables.iDelRow, "");
                            dtSource.SetValue("U_SOL_BRANDNAME", GeneralVariables.iDelRow, "");
                            dtSource.SetValue("U_SOL_BRANDDISC", GeneralVariables.iDelRow, "0");
                        }
                        else if (mtxName == "mt_2")
                        {
                            dtSource.SetValue("U_SOL_ITEMCODE", GeneralVariables.iDelRow, "");
                            dtSource.SetValue("U_SOL_ITEMNAME", GeneralVariables.iDelRow, "");
                            dtSource.SetValue("U_SOL_ITEMDISC", GeneralVariables.iDelRow, "0");
                        }

                        oMtx.LoadFromDataSource();

                        for (int i = 1; i <= oMtx.RowCount; i++)
                        {
                            oMtx.Columns.Item("#").Cells.Item(i).Specific.value = i;
                        }
                        oMtx.FlushToDataSource();

                        if (mtxName == "mt_2")
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
        /// Delete row in Fix Discount
        /// </summary>
        public void MenuEvent_FixDiscDel(ref Application oSBOApplication, ref MenuEvent pVal, ref bool bubbleEvent)
        {
            if (pVal.BeforeAction == true)
            {
                Form oForm = oSBOApplication.Forms.ActiveForm;
                if (oForm.TypeEx == "FIXDISC")
                {
                    try
                    {
                        oForm.Freeze(true);

                        string mtxName = string.Empty;
                        string dataSource = string.Empty;

                        switch (oForm.ActiveItem)
                        {
                            case "mt_1": mtxName = "mt_1"; dataSource = "@SOL_FIXDISC_D1"; break;
                            case "mt_2": mtxName = "mt_2"; dataSource = "@SOL_FIXDISC_D2"; break;
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

        #endregion

        #region Item Event
        /// <summary>
        /// Item Event Fix Discount
        /// </summary>
        public void ItemEvent_PrdDisc(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication,
                                        string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            switch (pVal.EventType)
            {
                case BoEventTypes.et_FORM_LOAD:; break;
                case BoEventTypes.et_CHOOSE_FROM_LIST: CFL_FixDisc(ref oSBOCompany, ref oSBOApplication, formUID, ref pVal, ref bubbleEvent); break;
                case BoEventTypes.et_VALIDATE: Validate_FixDisc(ref oSBOCompany, ref oSBOApplication, formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        /// <summary>
        /// Choose From List
        /// </summary>
        private void CFL_FixDisc(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication, string formUID, ref ItemEvent pVal,
                                        ref bool bubbleEvent)
        {
            IChooseFromListEvent oCFLEvent = null;
            try
            {
                oCFLEvent = (IChooseFromListEvent)pVal;

                switch (pVal.ItemUID)
                {
                    case "tCustCd": CFL_FixDisc_CustCode(ref oSBOCompany, ref oSBOApplication, formUID, ref pVal, ref bubbleEvent, ref oCFLEvent); break;
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
        /// CFL Customer Code
        /// </summary>
        private void CFL_FixDisc_CustCode(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication, string formUID, ref ItemEvent pVal,
                                        ref bool bubbleEvent, ref IChooseFromListEvent oCFLEvent)
        {
            if (bubbleEvent)
            {
                Form oForm = null;
                Conditions oCons = null;
                ICondition oCon = null;
                SAPbouiCOM.ChooseFromList oCFL = null;
                DataTable oDataTable = null;
                BusinessPartners oBp = null;
                DBDataSource oDBSource_H = null;

                try
                {
                    oForm = oSBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                    oBp = oSBOCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                    if (oCFLEvent.BeforeAction == true)
                    {
                        if (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_OK_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                        {
                            oForm.Freeze(true);
                            oCFL = oForm.ChooseFromLists.Item("CFL_CUST");

                            oCFL.SetConditions(null);
                            oCons = oCFL.GetConditions();

                            oCon = oCons.Add();
                            oCon.Alias = "CardType";
                            oCon.Operation = BoConditionOperation.co_EQUAL;
                            oCon.Relationship = BoConditionRelationship.cr_AND;
                            oCon.CondVal = "C";

                            oCon = oCons.Add();
                            oCon.Alias = "validFor";
                            oCon.Operation = BoConditionOperation.co_EQUAL;
                            oCon.CondVal = "Y";

                            oCFL.SetConditions(oCons);
                        }
                    }
                    else if (oCFLEvent.BeforeAction == false && oCFLEvent.ActionSuccess == true && oCFLEvent.SelectedObjects != null)
                    {
                        oForm.Freeze(true);
                        oDataTable = oCFLEvent.SelectedObjects;

                        if (oBp.GetByKey(oDataTable.GetValue("CardCode", 0)))
                        {
                            oDBSource_H = oForm.DataSources.DBDataSources.Item("@SOL_FIXDISC_H");
                            oDBSource_H.SetValue("U_SOL_CARDCODE", 0, oBp.CardCode);
                            oDBSource_H.SetValue("U_SOL_CARDNAME", 0, oBp.CardName);

                            // Generate Code
                            string code = GenerateCode(ref oSBOCompany, ref oSBOApplication, oBp.CardCode);
                            oDBSource_H.SetValue("Code", 0, code);

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
                    Utils.releaseObject(oBp);
                    Utils.releaseObject(oDBSource_H);
                }
            }
        }
        #endregion

        #region Right Click Event
        /// <summary>
        /// Create menu when Rihgt Click Event
        /// menu Add Row and Delete Row
        /// </summary>
        public void RightClickEvent_FixDisc(ref Application oSBOApplication, ref ContextMenuInfo eventInfo, ref bool bubbleEvent)
        {
            Form oForm = oSBOApplication.Forms.ActiveForm;

            if (eventInfo.BeforeAction == true && (eventInfo.ItemUID == "mt_1" || eventInfo.ItemUID == "mt_2"))
            {
                MenuItem oMenuItem = null;
                Menus oMenus = null;
                MenuCreationParams oCreateionPackage = null;

                try
                {
                    oCreateionPackage = oSBOApplication.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

                    oCreateionPackage.Type = BoMenuType.mt_STRING;
                    oCreateionPackage.UniqueID = "FixDiscAdd";
                    oCreateionPackage.Position = 1;
                    oCreateionPackage.String = "Add Row";
                    oCreateionPackage.Enabled = true;

                    oMenuItem = oSBOApplication.Menus.Item("1280");
                    oMenus = oMenuItem.SubMenus;
                    oMenus.AddEx(oCreateionPackage);

                    oCreateionPackage = oSBOApplication.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

                    oCreateionPackage.Type = BoMenuType.mt_STRING;
                    oCreateionPackage.UniqueID = "FixDiscDel";
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
                oSBOApplication.Menus.RemoveEx("FixDiscAdd");
                oSBOApplication.Menus.RemoveEx("FixDiscDel");
            }

            Utils.releaseObject(oForm);
        }
        #endregion

        #region Validate
        /// <summary>
        /// Validate Event
        /// </summary>
        private void Validate_FixDisc(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication, string formUID, ref ItemEvent pVal,
                                        ref bool bubbleEvent)
        {
            switch (pVal.ColUID)
            {
                case "cBrCd": Validate_FixDisc_BrandCode(ref oSBOCompany, ref oSBOApplication, formUID, ref pVal, ref bubbleEvent); break;
                case "cItmCd": Validate_FixDisc_ItemCode(ref oSBOCompany, ref oSBOApplication, formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        /// <summary>
        /// Validate - Brand Code
        /// Ketika Brand Code Dipilih, Brand Name terisi otomatis
        /// </summary>
        private void Validate_FixDisc_BrandCode(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication, string formUID, ref ItemEvent pVal,
                                        ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    Matrix oMtx = oForm.Items.Item("mt_1").Specific;
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_FIXDISC_D1");
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
                        string brandCode = oMtx.Columns.Item("cBrCd").Cells.Item(pVal.Row).Specific.Value;
                        string query = "SELECT T1.\"U_SOL_SUBGRP\" "
                                        + "FROM \"@SOL_ITMGRP_MAP_H\" T0 "
                                        + "INNER JOIN \"@SOL_ITMGRP_MAP_D\" T1 ON T0.\"Code\" = T1.\"Code\" "
                                        + "WHERE T1.\"U_SOL_SUBGRP_CODE\" = '" + brandCode + "'";

                        oRec.DoQuery(query);
                        if (oRec.RecordCount > 0)
                        {
                            nama = oRec.Fields.Item(0).Value;
                        }

                        dtSource.SetValue("U_SOL_BRANDNAME", pVal.Row - 1, nama);

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
        /// Validate - Item Code
        /// Ketika Item Code Dipilih, Item Name terisi otomatis
        /// </summary>
        private void Validate_FixDisc_ItemCode(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication, string formUID, ref ItemEvent pVal,
                                        ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    Matrix oMtx = oForm.Items.Item("mt_2").Specific;
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_FIXDISC_D2");
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


        #endregion

        /// <summary>
        /// Generate Code Fix Discount Master Data
        /// </summary>
        private string GenerateCode(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication,
                                    string custCode)
        {
            string runNumber = string.Empty;
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                    oRec.DoQuery("CALL SOL_SP_ADDON_FD_CODE ('" + custCode + "')");
                else
                    oRec.DoQuery("EXEC SOL_SP_ADDON_FD_CODE @CUSTCODE = '" + custCode + "'");

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
    }
}
