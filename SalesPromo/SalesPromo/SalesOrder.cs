using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;

namespace SalesPromo
{
    public class SalesOrder
    {
        public void ItemEvents_SalesOrder(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication,
                                            string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            #region Add Layout
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.EventType == BoEventTypes.et_FORM_LOAD)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    Item oItmBtn = null;
                    Button oBtn = null;
                    Item sapCancelButton = oForm.Items.Item("2");

                    try
                    {
                        oItmBtn = oForm.Items.Add("btnDisc", BoFormItemTypes.it_BUTTON);
                        oItmBtn.Top = sapCancelButton.Top;
                        oItmBtn.Left = sapCancelButton.Left + 70;
                        oItmBtn.Width = 110;
                        oItmBtn.Height = sapCancelButton.Height;
                        oItmBtn.DisplayDesc = true;
                        oItmBtn.Enabled = true;
                        oItmBtn.Visible = true;
                        oItmBtn.LinkTo = "2";
                        oBtn = oForm.Items.Item("btnDisc").Specific;
                        oBtn.Caption = "Calculate Discount";
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
                        Utils.releaseObject(oItmBtn);
                        Utils.releaseObject(oBtn);
                        Utils.releaseObject(sapCancelButton);
                    }
                }
            }
            #endregion

            #region Calculate Button
            if (bubbleEvent)
            {
                if (pVal.ItemUID == "btnDisc")
                {
                    if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.EventType == BoEventTypes.et_CLICK)
                    {
                        Form oForm = oSBOApplication.Forms.Item(formUID);
                        Form oUdfForm = oSBOApplication.Forms.Item(oForm.UDFFormUID);
                        Matrix oMtx = oForm.Items.Item("38").Specific;
                        Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                        try
                        {
                            GeneralVariables.oSOCurrent = new SOForm();
                            GeneralVariables.oSOCurrent.UniqueID = oForm.UniqueID;
                            GeneralVariables.oSOCurrent.CardCode = oForm.Items.Item("4").Specific.Value;
                            GeneralVariables.oSOCurrent.PostingDate = oForm.Items.Item("10").Specific.Value;

                            if (pVal.FormMode == (int)BoFormMode.fm_ADD_MODE)
                            {
                                oForm.Freeze(true);

                                ActiveRowUdf(true, ref oMtx, ref oUdfForm);
                                CalculateDiscount(ref oSBOCompany, ref oSBOApplication, oForm.UniqueID
                                                   , GeneralVariables.oSOCurrent.CardCode, GeneralVariables.oSOCurrent.PostingDate);
                            }
                        }
                        catch (Exception ex)
                        {
                            bubbleEvent = false;
                            oSBOApplication.MessageBox(ex.Message + " : " + ex.StackTrace);
                        }
                        finally
                        {
                            ActiveRowUdf(false, ref oMtx, ref oUdfForm);
                            if (oForm != null) oForm.Freeze(false);

                            Utils.releaseObject(oForm);
                            Utils.releaseObject(oMtx);
                            Utils.releaseObject(oRec);
                        }
                    }
                }
            }

            #endregion
        }

        private void SetColumnState(ref Application oSBOApplication, bool state)
        {
            Form oFormParent = oSBOApplication.Forms.Item(GeneralVariables.oSOCurrent.UniqueID);
            Matrix oMtxParent = oFormParent.Items.Item("38").Specific;

            try
            {
                if (state)
                {
                    oMtxParent.Columns.Item("1").Editable = true;
                    oMtxParent.Columns.Item("11").Editable = true;
                    oMtxParent.Columns.Item("1470002145").Editable = true;
                    oMtxParent.Columns.Item("U_SOL_ADDSC").Editable = true;
                    oMtxParent.Columns.Item("U_SOL_FD").Editable = true;
                    oMtxParent.Columns.Item("U_SOL_FDCD").Editable = true;
                    oMtxParent.Columns.Item("U_SOL_PD").Editable = true;
                    oMtxParent.Columns.Item("U_SOL_PDCD").Editable = true;
                    oMtxParent.Columns.Item("U_SOL_FLGBNS").Editable = true;
                }
                else
                {
                    oMtxParent.Columns.Item("1").Editable = GeneralVariables.oSOCurrent.ItemCodeState;
                    oMtxParent.Columns.Item("11").Editable = GeneralVariables.oSOCurrent.QtyState;
                    oMtxParent.Columns.Item("1470002145").Editable = GeneralVariables.oSOCurrent.UomState;
                    oMtxParent.Columns.Item("U_SOL_ADDSC").Editable = GeneralVariables.oSOCurrent.DiscAddonState;
                    oMtxParent.Columns.Item("U_SOL_FD").Editable = GeneralVariables.oSOCurrent.FixDiscState;
                    oMtxParent.Columns.Item("U_SOL_FDCD").Editable = false;
                    oMtxParent.Columns.Item("U_SOL_PD").Editable = GeneralVariables.oSOCurrent.PrdDiscState;
                    oMtxParent.Columns.Item("U_SOL_PDCD").Editable = false;
                    oMtxParent.Columns.Item("U_SOL_FLGBNS").Editable = false;
                }
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message + " : " + ex.StackTrace);
            }
            finally
            {
                Utils.releaseObject(oFormParent);
                Utils.releaseObject(oMtxParent);
            }
        }

        private void ClearAllDiscount(ref Application oSBOApplication)
        {
            Form oFormParent = oSBOApplication.Forms.Item(GeneralVariables.oSOCurrent.UniqueID);
            Matrix oMtxParent = oFormParent.Items.Item("38").Specific;
            ProgressBar oProgressBar = oSBOApplication.StatusBar.CreateProgressBar("Clear All Discount", oMtxParent.RowCount, true);
            oProgressBar.Text = "Clear all discount...";

            try
            {
                int i = 1;
                while (i != oMtxParent.RowCount)
                {
                    //if (oMtxParent.Columns.Item("U_SOL_BASEBONUS").Cells.Item(i).Specific.Value != "")
                    //{
                    //    oMtxParent.DeleteRow(i);
                    //}
                    //else
                    //{
                    oMtxParent.Columns.Item("U_SOL_FD").Cells.Item(i).Specific.Value = 0;
                    oMtxParent.Columns.Item("U_SOL_FDCD").Cells.Item(i).Specific.Value = "";
                    oMtxParent.Columns.Item("U_SOL_PD").Cells.Item(i).Specific.Value = 0;
                    oMtxParent.Columns.Item("U_SOL_PDCD").Cells.Item(i).Specific.Value = "";
                    i += 1;
                    //}
                }

                oProgressBar.Value += 1;
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message + " : " + ex.StackTrace);
            }
            finally
            {
                oProgressBar.Stop();
                Utils.releaseObject(oFormParent);
                Utils.releaseObject(oMtxParent);
                Utils.releaseObject(oProgressBar);
            }
        }

        /// <summary>
        /// Calculate discount
        /// </summary>
        private void CalculateDiscount(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication
                                        , string uniqueID, string cardCode, string postingDate)
        {
            Form oForm = oSBOApplication.Forms.Item(uniqueID);
            Form oUdfForm = oSBOApplication.Forms.Item(oForm.UDFFormUID);
            Matrix oMtx = oForm.Items.Item("38").Specific;
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            ProgressBar oProgressBar = oSBOApplication.StatusBar.CreateProgressBar("Calculate Discount", oMtx.RowCount, true);
            oProgressBar.Text = "Calculate Discount...";
            oForm.PaneLevel = 1;

            try
            {
                for (int i = 1; i < oMtx.RowCount; i++)
                {
                    string itemCode = oMtx.Columns.Item("1").Cells.Item(i).Specific.Value;
                    string address = oMtx.Columns.Item("275").Cells.Item(i).Specific.Value;
                    string area = GetAreaByCust(ref oSBOCompany, ref oSBOApplication, cardCode, address);
                    double qty = Utils.SBOToWindowsNumberWithoutCurrency(oMtx.Columns.Item("11").Cells.Item(i).Specific.Value);
                    double discount = Utils.SBOToWindowsNumberWithoutCurrency(oMtx.Columns.Item("15").Cells.Item(i).Specific.Value);

                    #region Fix Discount
                    string fixDiscCode = string.Empty;
                    double fixDisc = 0;

                    if (discount == 0 && oMtx.Columns.Item("40").Cells.Item(i).Specific.Value == "O" && oMtx.Columns.Item("U_SOL_FLGBNS").Cells.Item(i).Specific.Value != "Y")
                    {
                        string query = string.Empty;

                        if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                            query = "CALL SOL_SP_ADDON_GET_FIXDISC ('" + itemCode + "', '" + cardCode + "', '" + area + "')";
                        else
                            query = "EXEC SOL_SP_ADDON_GET_FIXDISC @ITEMCODE = '" + itemCode + "', @CUSTCODE = '" + cardCode + "', @AREA = '" + area + "'";

                        oRec.DoQuery(query);

                        if (oRec.RecordCount > 0)
                        {
                            fixDiscCode = oRec.Fields.Item("Code").Value;
                            fixDisc = Convert.ToDouble(Utils.FormattedStringAmount(oRec.Fields.Item("Disc").Value));

                            // Set value to row
                            oMtx.Columns.Item("U_SOL_FDCD").Cells.Item(i).Specific.Value = fixDiscCode;
                            oMtx.Columns.Item("U_SOL_FD").Cells.Item(i).Specific.Value = fixDisc;
                            string brandOrItem = oRec.Fields.Item("ItemOrBrandCode").Value;
                        }
                    }
                    #endregion

                    #region Periodic Discount
                    string prdDiscCode = string.Empty;
                    double prdDisc = 0;

                    if (discount == 0 && oMtx.Columns.Item("40").Cells.Item(i).Specific.Value == "O" && oMtx.Columns.Item("U_SOL_FLGBNS").Cells.Item(i).Specific.Value != "Y")
                    {
                        string query = string.Empty;

                        if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                            query = "CALL SOL_SP_ADDON_GET_PRDDISC ('" + itemCode + "', '" + cardCode + "', '" + area + "', '" + postingDate + "', '" + qty + "')";
                        else
                            query = "EXEC SOL_SP_ADDON_GET_PRDDISC @ITEMCODE = '" + itemCode + "', @CUSTCODE = '" + cardCode + "', @AREA = '" + area + "', @POSTINGDT = '" + postingDate + "', @QTY = '" + qty + "'";

                        oRec.DoQuery(query);

                        if (oRec.RecordCount > 0)
                        {
                            prdDiscCode = oRec.Fields.Item("Code").Value;

                            if (oRec.Fields.Item("DiscType").Value == "1")
                            {
                                if ((oRec.Fields.Item("CustType").Value == "All Customer" || oRec.Fields.Item("CustType").Value == "Per Customer") && oRec.Fields.Item("CustCode").Value == cardCode)
                                {
                                    prdDisc = Convert.ToDouble(Utils.SBOToWindowsNumberWithoutCurrency(oRec.Fields.Item("DiscPrcnt").Value));
                                    oMtx.Columns.Item("U_SOL_PDCD").Cells.Item(i).Specific.Value = prdDiscCode;
                                    oMtx.Columns.Item("U_SOL_PD").Cells.Item(i).Specific.Value = prdDisc;
                                }
                                else if (oRec.Fields.Item("CustType").Value == "All Customer")
                                {
                                    prdDisc = Convert.ToDouble(Utils.SBOToWindowsNumberWithoutCurrency(oRec.Fields.Item("DiscPrcnt").Value));

                                    oMtx.Columns.Item("U_SOL_PDCD").Cells.Item(i).Specific.Value = prdDiscCode;
                                    oMtx.Columns.Item("U_SOL_PD").Cells.Item(i).Specific.Value = prdDisc;
                                }
                            }
                            else
                            {
                                if ((oRec.Fields.Item("CustType").Value == "All Customer" || oRec.Fields.Item("CustType").Value == "Per Customer") && oRec.Fields.Item("CustCode").Value == cardCode)
                                {
                                    int currentRow = oMtx.RowCount;
                                    double qtyFree = Convert.ToDouble(Utils.SBOToWindowsNumberWithoutCurrency(oRec.Fields.Item("QtyFree").Value));
                                    double minQty = Convert.ToDouble(Utils.SBOToWindowsNumberWithoutCurrency(oRec.Fields.Item("MinqtyBG").Value));
                                    string kelipatan = oRec.Fields.Item("Kelipatan").Value;

                                    oMtx.Columns.Item("1").Cells.Item(currentRow).Specific.Value = oRec.Fields.Item("ItemCodeBG").Value;
                                    oMtx.Columns.Item("U_SOL_PDCD").Cells.Item(currentRow).Specific.Value = oRec.Fields.Item("Code").Value;
                                    oMtx.Columns.Item("U_SOL_FLGBNS").Cells.Item(currentRow).Specific.Value = "Y";

                                    if (kelipatan == "No")
                                        oMtx.Columns.Item("11").Cells.Item(currentRow).Specific.Value = qtyFree;
                                    else
                                        oMtx.Columns.Item("11").Cells.Item(currentRow).Specific.Value = Math.Round(qtyFree * (qty / minQty), 0);
                                }
                                else if (oRec.Fields.Item("CustType").Value == "All Customer")
                                {
                                    int currentRow = oMtx.RowCount;
                                    double qtyFree = Convert.ToDouble(Utils.SBOToWindowsNumberWithoutCurrency(oRec.Fields.Item("QtyFree").Value));
                                    double minQty = Convert.ToDouble(Utils.SBOToWindowsNumberWithoutCurrency(oRec.Fields.Item("MinqtyBG").Value));
                                    string kelipatan = oRec.Fields.Item("Kelipatan").Value;

                                    oMtx.Columns.Item("1").Cells.Item(currentRow).Specific.Value = oRec.Fields.Item("ItemCodeBG").Value;
                                    oMtx.Columns.Item("U_SOL_PDCD").Cells.Item(currentRow).Specific.Value = oRec.Fields.Item("Code").Value;
                                    oMtx.Columns.Item("U_SOL_FLGBNS").Cells.Item(currentRow).Specific.Value = "Y";

                                    if (kelipatan == "No")
                                        oMtx.Columns.Item("11").Cells.Item(currentRow).Specific.Value = qtyFree;
                                    else
                                        oMtx.Columns.Item("11").Cells.Item(currentRow).Specific.Value = Math.Round(qtyFree * (qty / minQty), 0);
                                }
                            }
                        }
                    }
                    #endregion

                    #region All Row Discount
                    if (oMtx.Columns.Item("U_SOL_FLGBNS").Cells.Item(i).Specific.Value == "Y")
                        prdDisc = 100;

                    double lineTotal = Utils.SBOToWindowsNumberWithCurrency(oMtx.Columns.Item("21").Cells.Item(i).Specific.Value);
                    double calculate = 0;
                    calculate = (lineTotal - (lineTotal * (fixDisc / 100))) - ((lineTotal - (lineTotal * (fixDisc / 100))) * (prdDisc / 100));

                    oMtx.Columns.Item("21").Cells.Item(i).Specific.Value = calculate;
                    oMtx.Columns.Item("U_SOL_ADDSC").Cells.Item(i).Specific.Value = (fixDisc + prdDisc) - ((fixDisc / 10) * (prdDisc / 10));
                    oMtx.Columns.Item("21").Cells.Item(i).Click();

                    #endregion
                }

                #region One Time Discount
                double discTotal = Utils.SBOToWindowsNumberWithoutCurrency(oForm.Items.Item("24").Specific.Value);
                string currency = oForm.Items.Item("63").Specific.Value;
                double price = Utils.SBOToWindowsNumberWithCurrency(oForm.Items.Item("22").Specific.Value);

                if (discTotal == 0)
                {
                    string query = string.Empty;

                    if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                        query = "CALL SOL_SP_ADDON_GET_CASHDISC ('" + cardCode + "', '" + currency + "', '" + price + "')";
                    else
                        query = "EXEC SOL_SP_ADDON_GET_CASHDISC @CUSTCODE = '" + cardCode + "', @CURR = '" + currency + "', @PRICE = '" + price + "'";

                    oRec.DoQuery(query);
                    if (oRec.RecordCount > 0)
                    {
                        oUdfForm.Items.Item("U_SOL_CASHDISC").Specific.Value = Convert.ToString(oRec.Fields.Item("U_SOL_CASHDISC").Value);
                        oForm.Items.Item("24").Specific.Value = Convert.ToString(oRec.Fields.Item("U_SOL_CASHDISC").Value);
                    }
                }

                #endregion

                // jika button diskon sudah di click. Untuk keperluan validasi
                oUdfForm.Items.Item("U_SOL_APPLDISC").Specific.Value = "Y";
                oUdfForm.Items.Item("U_SOL_MEMO").Click();
            }
            catch (Exception ex)
            {
                oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oProgressBar.Stop();
                if (oForm != null) oForm.Freeze(false);

                Utils.releaseObject(oForm);
                Utils.releaseObject(oMtx);
                Utils.releaseObject(oRec);
            }
        }

        /// <summary>
        /// Get Area(county) by Customer code and address code
        /// </summary>
        private string GetAreaByCust(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication,
                                        string cardCode, string address)
        {
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string area = string.Empty;

            try
            {
                string query = "SELECT T1.\"County\" "
                                + "FROM OCRD T0 LEFT JOIN CRD1 T1 ON T0.\"CardCode\" = T1.\"CardCode\" "
                                + "WHERE T1.\"AdresType\" = 'S' "
                                + "AND T1.\"Address\" = '" + address + "' "
                                + "AND T0.\"CardCode\" = '" + cardCode + "'";
                oRec.DoQuery(query);

                if (oRec.RecordCount > 0)
                {
                    area = oRec.Fields.Item("County").Value;
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

            return area;
        }

        /// <summary>
        /// enable and disable udf in Row
        /// </summary>
        private void ActiveRowUdf(bool param, ref Matrix oMtx, ref Form oUdfForm)
        {
            oMtx.Columns.Item("U_SOL_ADDSC").Editable = param;
            oMtx.Columns.Item("U_SOL_FD").Editable = param;
            oMtx.Columns.Item("U_SOL_PD").Editable = param;
            oMtx.Columns.Item("U_SOL_PDCD").Editable = param;
            oMtx.Columns.Item("U_SOL_FDCD").Editable = param;
            oMtx.Columns.Item("U_SOL_FLGBNS").Editable = param;
            oUdfForm.Items.Item("U_SOL_CASHDISC").Enabled = param;
            oUdfForm.Items.Item("U_SOL_APPLDISC").Enabled = param;
        }
    }
}
