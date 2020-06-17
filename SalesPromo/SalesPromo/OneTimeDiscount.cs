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
        /// <summary>
        /// Menu Event One Time Discount
        /// When click menu, this event called
        /// </summary>
        public void MenuEvent_CashDisc(ref Application oSBOApplication, ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if(pVal.BeforeAction == false)
            {
                Form oForm = null;

                try
                {
                    oForm = Utils.createForm(ref oSBOApplication, "CashDisc");
                    oForm.Visible = true;
                }
                catch(Exception ex)
                {
                    bubbleEvent = false;
                    oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                finally
                {
                    if(oForm != null)
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
        public void ItemEvent_CashDisc(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication, 
                                        string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            switch (pVal.EventType)
            {
                case BoEventTypes.et_FORM_LOAD:; break;
                case BoEventTypes.et_CHOOSE_FROM_LIST: CFL_CashDisc(ref oSBOCompany, ref oSBOApplication, formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        private void CFL_CashDisc(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication, string formUID, ref ItemEvent pVal,
                                        ref bool bubbleEvent)
        {
            IChooseFromListEvent oCFLEvent = null;
            try
            {
                oCFLEvent = (IChooseFromListEvent)pVal;

                switch (pVal.ItemUID)
                {
                    case "tCusCd": CFL_CashDisc_Cust(ref oSBOCompany, ref oSBOApplication, formUID, ref pVal, ref bubbleEvent, ref oCFLEvent); break;
                }

            }
            catch(Exception ex)
            {
                bubbleEvent = false;
                oSBOApplication.MessageBox(ex.Message);
            }
            finally
            {
                Utils.releaseObject(oCFLEvent);
            }
        }

        private void CFL_CashDisc_Cust(ref SAPbobsCOM.Company oSBOCompany,ref Application oSBOApplication, string formUID, ref ItemEvent pVal,
                                        ref bool bubbleEvent, ref IChooseFromListEvent oCFLEvent)
        {
            if (bubbleEvent)
            {
                Form oForm = null;
                Conditions oCons = null;
                Condition oCon = null;
                SAPbouiCOM.ChooseFromList oCFL = null;
                DataTable oDataTable = null;
                BusinessPartners oBP = null;
                DBDataSource oDBSource_H = null;

                try
                {
                    oForm = oSBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                    oBP = oSBOCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                    if (oCFLEvent.BeforeAction == true)
                    {
                        if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            oForm.Freeze(true);
                            oCFL = oForm.ChooseFromLists.Item("CFL_CUST");

                            // Add conditions
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
                        if (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE || oForm.Mode == BoFormMode.fm_OK_MODE)
                        {
                            oForm.Freeze(true);
                            oDataTable = oCFLEvent.SelectedObjects;

                            if (oBP.GetByKey(oDataTable.GetValue("CardCode", 0)))
                            {
                                oDBSource_H = oForm.DataSources.DBDataSources.Item("@SOL_CASHDISC");
                                oDBSource_H.SetValue("U_SOL_CARDCODE", 0, oBP.CardCode);

                                if (oForm.Mode != BoFormMode.fm_ADD_MODE)
                                    oForm.Mode = BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                    }
                }
                catch(Exception ex)
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
                    Utils.releaseObject(oBP);
                    Utils.releaseObject(oDBSource_H);
                }
            }
        }
        #endregion
    }
}
