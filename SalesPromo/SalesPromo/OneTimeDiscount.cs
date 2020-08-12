using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SalesPromo
{
    public class OneTimeDiscount
    {
        private SAPbouiCOM.Application oSBOApplication;
        private SAPbobsCOM.Company oSBOCompany;

        public OneTimeDiscount(SAPbouiCOM.Application oSBOApplication, SAPbobsCOM.Company oSBOCompany)
        {
            this.oSBOApplication = oSBOApplication;
            this.oSBOCompany = oSBOCompany;
        }

        /// <summary>
        /// Menu Event One Time Discount
        /// When click menu, this event called
        /// </summary>
        public void MenuEvent_CashDisc(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (pVal.BeforeAction == false)
            {
                Form oForm = null;

                try
                {
                    oForm = Utils.createForm(ref oSBOApplication, "CashDisc");
                    oForm.Visible = true;
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
        /// Item Event
        /// </summary>
        #region Item Event
        public void ItemEvent_CashDisc(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            switch (pVal.EventType)
            {
                case BoEventTypes.et_VALIDATE: Validate_CashDisc(formUID, ref pVal, ref bubbleEvent);break;
                case BoEventTypes.et_CLICK: CashDisc_Click(formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        /// <summary>
        /// Validate
        /// </summary>
        private void Validate_CashDisc(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            switch (pVal.ItemUID)
            {
                case "tCusCd": Validate_CashDisc_CustCode(formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        /// <summary>
        /// Reload no memo when click add
        /// </summary>
        private void CashDisc_Click(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            if(pVal.ItemUID == "1")
            {
                if(pVal.FormMode == (int)BoFormMode.fm_ADD_MODE)
                {
                    if(pVal.BeforeAction == true)
                    {
                        Form oForm = oSBOApplication.Forms.Item(formUID);
                        string cardCode = oForm.Items.Item("tCusCd").Specific.Value;
                        oForm.Items.Item("txtCode").Specific.Value = GenerateNoMemo(cardCode);
                    }
                }
            }
        }

        /// <summary>
        /// Validate customer code - nomor memo
        /// </summary>
        private void Validate_CashDisc_CustCode(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_CASHDISC");
                    Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    try
                    {
                        oForm.Freeze(true);
                        string noMemo = GenerateNoMemo(oForm.Items.Item("tCusCd").Specific.Value);
                        dtSource.SetValue("Code", 0, noMemo);
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
                        Utils.releaseObject(dtSource);
                        Utils.releaseObject(oRec);
                    }
                }
            }
        }
        #endregion

        /// <summary>
        /// Generate nomor memo
        /// </summary>
        private string GenerateNoMemo(string cardCode)
        {
            string runNumber = string.Empty;
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                    oRec.DoQuery("CALL SOL_SP_ADDON_CD_CODE ('" + cardCode + "')");

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
