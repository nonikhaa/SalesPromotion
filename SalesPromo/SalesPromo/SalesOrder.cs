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
        private SAPbouiCOM.Application oSBOApplication;
        private SAPbobsCOM.Company oSBOCompany;

        public SalesOrder(SAPbouiCOM.Application oSBOApplication, SAPbobsCOM.Company oSBOCompany)
        {
            this.oSBOApplication = oSBOApplication;
            this.oSBOCompany = oSBOCompany;
        }

        public void ItemEvents_SalesOrder(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            #region Add Layout
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.EventType == BoEventTypes.et_FORM_LOAD)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    //Matrix oMtx = oForm.Items.Item("38").Specific;
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

                        //EditText oFlagBns = oMtx.Columns.Item("U_SOL_FLGBNS").Cells.Item(0).Specific;
                        //oFlagBns.Active = true;
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
                        string docNum = oForm.Items.Item("8").Specific.Value;
                        Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                        try
                        {
                            GeneralVariables.oSOCurrent = new SOForm();
                            GeneralVariables.oSOCurrent.UniqueID = oForm.UniqueID;
                            GeneralVariables.oSOCurrent.CardCode = oForm.Items.Item("4").Specific.Value;
                            GeneralVariables.oSOCurrent.PostingDate = oForm.Items.Item("10").Specific.Value;
                            string tipeSO = oUdfForm.Items.Item("U_SOL_TIPE_SO").Specific.Value;
                            string errorMsg = string.Empty;

                            // validate
                            errorMsg = AllowedCalculate(oMtx, GeneralVariables.oSOCurrent.CardCode, tipeSO, docNum);

                            if (string.IsNullOrEmpty(errorMsg))
                            {
                                if (pVal.FormMode == (int)BoFormMode.fm_ADD_MODE || pVal.FormMode == (int)BoFormMode.fm_UPDATE_MODE)
                                {
                                    oForm.Freeze(true);

                                    ActiveRowUdf(true, ref oMtx, ref oUdfForm);
                                    CalculateDiscount(oForm.UniqueID, GeneralVariables.oSOCurrent.CardCode, GeneralVariables.oSOCurrent.PostingDate);
                                }
                            }
                            else
                            {
                                oSBOApplication.MessageBox(errorMsg);
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
        private void CalculateDiscount(string uniqueID, string cardCode, string postingDate)
        {
            Form oForm = oSBOApplication.Forms.Item(uniqueID);
            Form oUdfForm = oSBOApplication.Forms.Item(oForm.UDFFormUID);
            Matrix oMtx = oForm.Items.Item("38").Specific;

            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            ClearPromo(ref oMtx, ref oForm, ref oUdfForm);

            ProgressBar oProgressBar = oSBOApplication.StatusBar.CreateProgressBar("Calculate Discount", oMtx.RowCount, true);
            oProgressBar.Text = "Calculate Discount...";
            oForm.PaneLevel = 1;

            try
            {
                List<MatrixSo> GroupFixDisc = null;
                List<MatrixSo> GroupPrdDisc = null;
                List<OutputDiscQuery> listDiscSO = new List<OutputDiscQuery>();

                GroupingItem(ref oMtx, cardCode, out GroupFixDisc, out GroupPrdDisc);
                GetDiscount(ref oMtx, cardCode, postingDate, out listDiscSO, GroupFixDisc, GroupPrdDisc);
                ApplyDiscount(ref oMtx, oForm, oUdfForm, cardCode, listDiscSO, GroupFixDisc, GroupPrdDisc);
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
        /// Grouping Item
        /// </summary>
        private void GroupingItem(ref Matrix oMtx, string cardCode, out List<MatrixSo> groupFixDisc, out List<MatrixSo> groupPrdDisc)
        {
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            groupFixDisc = new List<MatrixSo>();
            groupPrdDisc = new List<MatrixSo>();

            for (int i = 1; i < oMtx.RowCount; i++)
            {
                string itemCode = oMtx.Columns.Item("1").Cells.Item(i).Specific.Value;
                string qtySAP = oMtx.Columns.Item("11").Cells.Item(i).Specific.Value;
                double qty = Convert.ToDouble(qtySAP.Replace(".", ","));
                string address = oMtx.Columns.Item("275").Cells.Item(i).Specific.Value;
                string area = GetAreaByCust(cardCode, address);
                string detailStatus = oMtx.Columns.Item("40").Cells.Item(i).Specific.Value;
                string itemBonus = oMtx.Columns.Item("U_SOL_FLGBNS").Cells.Item(i).Specific.Value;
                string delDate = oMtx.Columns.Item("25").Cells.Item(i).Specific.Value;

                #region Fix Disc

                // check discount ALL area
                string query = string.Empty;
                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                    query = "CALL SOL_SP_ADDON_GET_DISCAREA ('" + itemCode + "', '" + cardCode + "')";
                else
                    query = "EXEC SOL_SP_ADDON_GET_DISCAREA @ITEMCODE = '" + itemCode + "', @CUSTCODE = '" + cardCode + "'";

                oRec.DoQuery(query);

                MatrixSo mtxFixDisc = new MatrixSo();
                if (oRec.RecordCount > 0 && groupFixDisc.Select(o => o.ItemCode).ToList().Contains(itemCode))
                {
                    groupFixDisc.Where(o => o.ItemCode == itemCode).ToList().ForEach(a =>
                    {
                        a.Quantity += qty;
                        a.Area = "ALL";
                        if (DateTime.ParseExact(a.DeliveryDate, "yyyyMMdd", null) < DateTime.ParseExact(delDate, "yyyyMMdd", null))
                            a.DeliveryDate = delDate;
                    });
                }
                else if (oRec.RecordCount <= 0 && groupFixDisc.Where(o => o.ItemCode == itemCode && o.Area == area).Select(o => o.ItemCode).ToList().Contains(itemCode))
                {
                    groupFixDisc.Where(o => o.ItemCode == itemCode && o.Area == area).ToList().ForEach(a =>
                    {
                        a.Quantity += qty;
                        a.Area = area;
                        if (DateTime.ParseExact(a.DeliveryDate, "yyyyMMdd", null) < DateTime.ParseExact(delDate, "yyyyMMdd", null))
                            a.DeliveryDate = delDate;
                    });
                }
                else
                {
                    mtxFixDisc.ItemCode = itemCode;
                    mtxFixDisc.Quantity = qty;
                    mtxFixDisc.FlagBonus = itemBonus;
                    mtxFixDisc.Area = area;
                    mtxFixDisc.DeliveryDate = delDate;

                    groupFixDisc.Add(mtxFixDisc);
                }

                #endregion

                #region Prd Disc
                MatrixSo mtxPrdDisc = new MatrixSo();
                if (groupPrdDisc.Select(o => o.ItemCode).ToList().Contains(itemCode))
                {
                    groupPrdDisc.Where(o => o.ItemCode == itemCode).ToList().ForEach(a =>
                    {
                        a.Quantity += qty;
                        if (DateTime.ParseExact(a.DeliveryDate, "yyyyMMdd", null) < DateTime.ParseExact(delDate, "yyyyMMdd", null))
                            a.DeliveryDate = delDate;
                    });
                }
                else
                {
                    mtxPrdDisc.ItemCode = itemCode;
                    mtxPrdDisc.Quantity = qty;
                    mtxPrdDisc.Area = area;
                    mtxPrdDisc.FlagBonus = itemBonus;
                    mtxPrdDisc.DeliveryDate = delDate;

                    groupPrdDisc.Add(mtxPrdDisc);
                }
                #endregion
            }
        }


        /// <summary>
        /// Get Discount
        /// </summary>
        private void GetDiscount(ref Matrix oMtx, string cardCode, string postingDate, out List<OutputDiscQuery> listDiscSO
            , List<MatrixSo> groupFixDisc, List<MatrixSo> groupPrdDisc)
        {
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            listDiscSO = new List<OutputDiscQuery>();

            #region Fix Discount
            foreach (var detailFix in groupFixDisc)
            {
                OutputDiscQuery discSO = new OutputDiscQuery();

                if (detailFix.FlagBonus != "Y")
                {
                    string query = string.Empty;

                    if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                        query = "CALL SOL_SP_ADDON_GET_FIXDISC ('" + detailFix.ItemCode + "', '" + cardCode + "', '" + detailFix.Area + "')";
                    else
                        query = "EXEC SOL_SP_ADDON_GET_FIXDISC @ITEMCODE = '" + detailFix.ItemCode + "', @CUSTCODE = '" + cardCode + "', @AREA = '" + detailFix.Area + "'";

                    oRec.DoQuery(query);
                    if (oRec.RecordCount > 0)
                    {
                        discSO.ItemCode = detailFix.ItemCode;
                        discSO.FixDiscCode = oRec.Fields.Item("Code").Value;
                        discSO.FixDisc = Convert.ToDouble(Utils.FormattedStringAmount(oRec.Fields.Item("Disc").Value));
                        discSO.Area = detailFix.Area;
                        listDiscSO.Add(discSO);
                    }
                }
            }
            #endregion

            #region Periodic Discount
            foreach (var detailPrd in groupPrdDisc)
            {
                OutputDiscQuery discSO = new OutputDiscQuery();

                if (detailPrd.FlagBonus != "Y")
                {
                    string query = string.Empty;

                    if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                        query = "CALL SOL_SP_ADDON_GET_PRDDISC ('" + detailPrd.ItemCode + "', '" + cardCode + "', '" + postingDate + "', '" + detailPrd.Quantity + "')";
                    else
                        query = "EXEC SOL_SP_ADDON_GET_PRDDISC @ITEMCODE = '" + detailPrd.ItemCode + "', @CUSTCODE = '" + cardCode + "', @POSTINGDT = '" + postingDate + "', @QTY = '" + detailPrd.Quantity + "'";

                    oRec.DoQuery(query);
                    if (oRec.RecordCount > 0)
                    {
                        if (listDiscSO.Select(o => o.ItemCode).ToList().Contains(detailPrd.ItemCode))
                        {
                            listDiscSO.Where(o => o.ItemCode == detailPrd.ItemCode).ToList()
                                .ForEach(a =>
                                {
                                    a.DiscountType = oRec.Fields.Item("DiscType").Value;
                                    a.CustomerType = oRec.Fields.Item("CustType").Value;
                                    a.CustomerCode = oRec.Fields.Item("CustCode").Value;
                                    a.PrdDiscCode = oRec.Fields.Item("Code").Value;
                                    a.PrcntDisc = oRec.Fields.Item("DiscPrcnt").Value;
                                    a.PrcntMinQty = oRec.Fields.Item("MinQtyPrcnt").Value;
                                    a.BXGYItemCd = oRec.Fields.Item("ItemCodeBG").Value;
                                    a.BXGYMinQty = oRec.Fields.Item("MinQtyBG").Value;
                                    a.BXGYItemBns = oRec.Fields.Item("ItemCodeFree").Value;
                                    a.BXGYQtyFree = oRec.Fields.Item("QtyFree").Value;
                                    a.Kelipatan = oRec.Fields.Item("Kelipatan").Value;
                                });
                        }
                        else
                        {
                            discSO.ItemCode = oRec.Fields.Item("ItemCodePrcnt").Value;
                            discSO.DiscountType = oRec.Fields.Item("DiscType").Value;
                            discSO.CustomerType = oRec.Fields.Item("CustType").Value;
                            discSO.CustomerCode = oRec.Fields.Item("CustCode").Value;
                            discSO.PrdDiscCode = oRec.Fields.Item("Code").Value;
                            discSO.PrcntDisc = oRec.Fields.Item("DiscPrcnt").Value;
                            discSO.PrcntMinQty = oRec.Fields.Item("MinQtyPrcnt").Value;
                            discSO.BXGYItemCd = oRec.Fields.Item("ItemCodeBG").Value;
                            discSO.BXGYMinQty = oRec.Fields.Item("MinQtyBG").Value;
                            discSO.BXGYItemBns = oRec.Fields.Item("ItemCodeFree").Value;
                            discSO.BXGYQtyFree = oRec.Fields.Item("QtyFree").Value;
                            discSO.Kelipatan = oRec.Fields.Item("Kelipatan").Value;
                            listDiscSO.Add(discSO);
                        }
                    }
                }
            }
            #endregion
        }


        /// <summary>
        /// Apply Discount
        /// </summary>
        private void ApplyDiscount(ref Matrix oMtx, Form oForm, Form oUdfForm, string cardCode, List<OutputDiscQuery> listDiscSO
            , List<MatrixSo> groupFixDisc, List<MatrixSo> groupPrdDisc)
        {
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            for (int i = 1; i < oMtx.RowCount; i++)
            {
                string itemCode = oMtx.Columns.Item("1").Cells.Item(i).Specific.Value;
                string qtySAP = oMtx.Columns.Item("11").Cells.Item(i).Specific.Value;
                double qty = Convert.ToDouble(qtySAP.Replace(".", ","));
                string curr = oForm.Items.Item("63").Specific.Value;
                string address = oMtx.Columns.Item("275").Cells.Item(i).Specific.Value;
                string area = GetAreaByCust(cardCode, address);

                var dataDiscount = listDiscSO.Where(o => o.ItemCode == itemCode);

                double prdDisc = 0;
                double fixDisc = 0;

                #region FixDiscount
                if (dataDiscount.Count() > 0 && (dataDiscount.FirstOrDefault().Area == area || dataDiscount.FirstOrDefault().Area == "ALL"))
                {
                    oMtx.Columns.Item("U_SOL_FDCD").Cells.Item(i).Specific.Value = dataDiscount.FirstOrDefault().FixDiscCode;
                    oMtx.Columns.Item("U_SOL_FD").Cells.Item(i).Specific.Value = Utils.WindowsToSBONumber(dataDiscount.FirstOrDefault().FixDisc);
                    fixDisc = dataDiscount.FirstOrDefault().FixDisc;
                }
                #endregion

                #region Periodic Discount - Discount %
                if (dataDiscount.Count() > 0)
                {
                    // by discount type
                    if (dataDiscount.FirstOrDefault().DiscountType == "1")
                    {
                        // by customer type
                        if (dataDiscount.FirstOrDefault().CustomerType == "All Customer" || (dataDiscount.FirstOrDefault().CustomerType == "Per Customer" && dataDiscount.FirstOrDefault().CustomerCode == cardCode))
                        {
                            oMtx.Columns.Item("U_SOL_PDCD").Cells.Item(i).Specific.Value = dataDiscount.FirstOrDefault().PrdDiscCode;
                            oMtx.Columns.Item("U_SOL_PD").Cells.Item(i).Specific.Value = Utils.WindowsToSBONumber(dataDiscount.FirstOrDefault().PrcntDisc);
                            prdDisc = dataDiscount.FirstOrDefault().PrcntDisc;
                        }
                    }
                }
                #endregion

                #region All Row Discount
                double lineTotal = 0;

                // Ambil line total
                if (curr == "IDR")
                    lineTotal = Utils.SBOToWindowsNumberWithCurrency(oMtx.Columns.Item("21").Cells.Item(i).Specific.Value);
                else
                    lineTotal = Utils.SBOToWindowsNumberWithCurrency(oMtx.Columns.Item("23").Cells.Item(i).Specific.Value);

                // Hitung diskon
                double calculate = 0;
                calculate = (lineTotal - (lineTotal * (fixDisc / 100))) - ((lineTotal - (lineTotal * (fixDisc / 100))) * (prdDisc / 100));

                // Assign jumlah line total
                if (curr == "IDR")
                    oMtx.Columns.Item("21").Cells.Item(i).Specific.Value = Utils.WindowsToSBONumber(calculate);
                else
                    oMtx.Columns.Item("23").Cells.Item(i).Specific.Value = Utils.WindowsToSBONumber(calculate);

                oMtx.Columns.Item("U_SOL_ADDSC").Cells.Item(i).Specific.Value = Utils.WindowsToSBONumber((fixDisc + prdDisc) - ((fixDisc / 10) * (prdDisc / 10)));
                oMtx.Columns.Item("3").Cells.Item(i).Click();

                #endregion
            }

            #region Periodic Discount - Buy X get Y
            var dataBonus = listDiscSO.Where(o => o.DiscountType == "2").ToList();
            string tipeSO = oUdfForm.Items.Item("U_SOL_TIPE_SO").Specific.Value;

            if (dataBonus.Count() > 0)
            {
                foreach (var detail in dataBonus)
                {
                    if (detail.CustomerType == "All Customer" || (detail.CustomerType == "Per Customer" && detail.CustomerCode == cardCode))
                    {
                        var groupItem = groupPrdDisc.Where(o => o.ItemCode == detail.BXGYItemCd).FirstOrDefault();
                        int currentRow = oMtx.RowCount;
                        double qtyFree = detail.BXGYQtyFree;
                        double minQty = detail.BXGYMinQty;

                        oMtx.Columns.Item("1").Cells.Item(currentRow).Specific.Value = detail.BXGYItemBns;
                        oMtx.Columns.Item("25").Cells.Item(currentRow).Specific.Value = groupItem.DeliveryDate;
                        oMtx.Columns.Item("U_SOL_PDCD").Cells.Item(currentRow).Specific.Value = detail.PrdDiscCode;
                        oMtx.Columns.Item("U_SOL_FLGBNS").Cells.Item(currentRow).Specific.Value = "Y";
                        oMtx.Columns.Item("U_SOL_ADDSC").Cells.Item(currentRow).Specific.Value = 100;
                        oMtx.Columns.Item("U_SOL_ITMCDORI").Cells.Item(currentRow).Specific.Value = detail.ItemCode;
                        oMtx.Columns.Item("U_SOL_ITMNMORI").Cells.Item(currentRow).Specific.Value = GetItemName(detail.ItemCode);
                        double prc = GetPrice(detail.BXGYItemBns, cardCode, tipeSO, GetDefaultAddress(cardCode));
                        oMtx.Columns.Item("20").Cells.Item(currentRow).Specific.Value = prc; // gross price after disc
                        oMtx.Columns.Item("15").Cells.Item(currentRow).Specific.Value = 100; // discount

                        if (detail.Kelipatan == "N")
                            oMtx.Columns.Item("11").Cells.Item(currentRow).Specific.Value = qtyFree;
                        else
                        {
                            string jmlKelipatan = (qtyFree * (groupItem.Quantity / minQty)).ToString();
                            if (jmlKelipatan.IndexOf(',') > 0)
                                oMtx.Columns.Item("11").Cells.Item(currentRow).Specific.Value = jmlKelipatan.Substring(0, jmlKelipatan.IndexOf(","));
                            else
                                oMtx.Columns.Item("11").Cells.Item(currentRow).Specific.Value = jmlKelipatan;
                        }
                    }
                }
            }
            #endregion

            #region One Time Discount
            string disc = oForm.Items.Item("24").Specific.Value;
            double discTotal = Convert.ToDouble(disc.Replace(".", ","));
            string currency = oForm.Items.Item("63").Specific.Value;
            double priceA = Utils.SBOToWindowsNumberWithCurrency(oForm.Items.Item("22").Specific.Value);
            string price = priceA.ToString().Replace(",", ".");

            if (discTotal == 0)
            {
                string query = string.Empty;

                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                    query = "CALL SOL_SP_ADDON_GET_CASHDISC ('" + cardCode + "', '" + currency + "', '" + price + "')";
                else
                    query = "EXEC SOL_SP_ADDON_GET_CASHDISC @CUSTCODE = '" + cardCode + "', @CURR = '" + currency + "', @PRICE = " + price + "";

                oRec.DoQuery(query);
                if (oRec.RecordCount > 0)
                {
                    oUdfForm.Items.Item("U_SOL_CASHDISC").Specific.Value = Convert.ToString(oRec.Fields.Item("U_SOL_CASHDISC").Value);
                    oUdfForm.Items.Item("U_SOL_MEMO").Specific.Value = oRec.Fields.Item("Code").Value;
                    oForm.Items.Item("24").Specific.Value = Convert.ToString(oRec.Fields.Item("U_SOL_CASHDISC").Value);
                }
            }
            #endregion

            oUdfForm.Items.Item("U_SOL_APPLDISC").Specific.Value = "Y";
            oUdfForm.Items.Item("U_SOL_MEMO").Click();
        }


        /// <summary>
        /// Get Area(county) by Customer code and address code
        /// </summary>
        private string GetAreaByCust(string cardCode, string address)
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
            oMtx.Columns.Item("U_SOL_ITMCDORI").Editable = param;
            oMtx.Columns.Item("U_SOL_ITMNMORI").Editable = param;
            oUdfForm.Items.Item("U_SOL_CASHDISC").Enabled = param;
            oUdfForm.Items.Item("U_SOL_APPLDISC").Enabled = param;
        }

        /// <summary>
        /// Clear promo before calculate discount
        /// </summary>
        private void ClearPromo(ref Matrix oMtx, ref Form oForm, ref Form oUdfForm)
        {
            for (int i = 1; i < oMtx.RowCount; i++)
            {
                if (oMtx.Columns.Item("U_SOL_FLGBNS").Cells.Item(i).Specific.Value == "Y")
                {
                    oMtx.DeleteRow(i);
                }
                else
                {
                    oMtx.Columns.Item("15").Cells.Item(i).Specific.Value = 0; //discount
                    oMtx.Columns.Item("U_SOL_ADDSC").Cells.Item(i).Specific.Value = 0; // Addon disc
                    oMtx.Columns.Item("U_SOL_PDCD").Cells.Item(i).Specific.Value = ""; // Periodic code
                    oMtx.Columns.Item("U_SOL_PD").Cells.Item(i).Specific.Value = 0; // Periodic discount
                    oMtx.Columns.Item("U_SOL_FDCD").Cells.Item(i).Specific.Value = ""; // Fix discount code
                    oMtx.Columns.Item("U_SOL_FD").Cells.Item(i).Specific.Value = 0; // Fix discount
                    oUdfForm.Items.Item("U_SOL_CASHDISC").Specific.Value = 0;
                    oUdfForm.Items.Item("U_SOL_MEMO").Specific.Value = "";
                    oForm.Items.Item("24").Specific.Value = 0; // discount header
                }
            }
        }

        /// <summary>
        /// Check user group allowed for calculate or not
        /// </summary>
        private string AllowedCalculate(Matrix oMtx, string cardCode, string tipeSO, string DocNum)
        {
            string message = string.Empty;
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            for (int i = 1; i < oMtx.RowCount; i++)
            {
                string itemCode = oMtx.Columns.Item("1").Cells.Item(i).Specific.Value;
                string shipTo = oMtx.Columns.Item("275").Cells.Item(i).Specific.Value;

                string query = "CALL SOL_SP_ADDON_VALIDATE_CALCULATE('" + cardCode + "', '" + tipeSO + "', '" + DocNum + "', '" + itemCode + "', '" + shipTo + "')";
                oRec.DoQuery(query);

                if (oRec.RecordCount > 0)
                {
                    message = oRec.Fields.Item("Message").Value;
                    break;
                }
            }

            Utils.releaseObject(oRec);
            return message;
        }
               
        /// <summary>
        /// Get Item Name
        /// </summary>
        private string GetItemName(string itemCode)
        {
            string itemName = string.Empty;
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "SELECT \"ItemName\" FROM OITM WHERE \"ItemCode\" = '" + itemCode + "'";
            oRec.DoQuery(query);

            if (oRec.RecordCount > 0)
            {
                itemName = oRec.Fields.Item("ItemName").Value;
            }

            Utils.releaseObject(oRec);
            return itemName;
        }

        /// <summary>
        /// Get gross price after discount
        /// </summary>
        private double GetPrice(string itemCode, string cardCode, string tipeSO, string shipTo)
        {
            double price = 0;
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "CALL SOL_SP_ADDON_GET_PRICE('" + itemCode + "', '" + cardCode + "', '" + tipeSO + "', '" + shipTo + "')";
            oRec.DoQuery(query);

            if (oRec.RecordCount > 0)
            {
                price = oRec.Fields.Item("Price").Value;
            }

            Utils.releaseObject(oRec);
            return price;
        }

        /// <summary>
        /// Get default ship to address
        /// </summary>
        private string GetDefaultAddress(string cardCode)
        {
            string address = string.Empty;
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "SELECT OCRD.\"ShipToDef\" FROM OCRD "
                            + "LEFT JOIN CRD1 ON OCRD.\"CardCode\" = CRD1.\"CardCode\" "
                            + "WHERE OCRD.\"CardCode\" = '" + cardCode + "' "
                            + "AND CRD1.\"AdresType\" = 'S'";

            oRec.DoQuery(query);

            if (oRec.RecordCount > 0)
            {
                address = oRec.Fields.Item("ShipToDef").Value;
            }

            Utils.releaseObject(oRec);
            return address;

        }
    }
}
