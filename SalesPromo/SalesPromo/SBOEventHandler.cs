using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SalesPromo
{
    public class SBOEventHandler
    {
        private SAPbouiCOM.Application oSBOApplication;
        private SAPbobsCOM.Company oSBOCompany;
        public SAPbobsCOM.CompanyService oCompService;

        /// <summary>
        /// Constructor --> first initialization when class is called
        /// </summary>
        #region Constructor
        public SBOEventHandler()
        {

        }

        public SBOEventHandler(SAPbouiCOM.Application oSBOApplication)
        {
            this.oSBOApplication = oSBOApplication;
        }

        public SBOEventHandler(SAPbouiCOM.Application oSBOApplication, SAPbobsCOM.Company oSBOCompany)
        {
            this.oSBOApplication = oSBOApplication;
            this.oSBOCompany = oSBOCompany;
        }
        #endregion

        /// <summary>
        /// Handle App Event
        /// </summary>
        public void HandleAppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    if (oSBOCompany.Connected) oSBOCompany.Disconnect();
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    if (oSBOCompany.Connected) oSBOCompany.Disconnect();
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    if (oSBOCompany.Connected) oSBOCompany.Disconnect();
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    if (oSBOCompany.Connected) oSBOCompany.Disconnect();
                    System.Windows.Forms.Application.Exit();
                    break;
            }
        }

        /// <summary>
        /// Handle Menu Event
        /// </summary>
        public void HandleMenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            OneTimeDiscount oneTimeDisc = new OneTimeDiscount();
            PeriodicDiscount prdDisc = new PeriodicDiscount();
            FixDiscount fixDisc = new FixDiscount();

            try
            {
                switch (pVal.MenuUID)
                {
                    case "96969": oneTimeDisc.MenuEvent_CashDisc(ref oSBOApplication, ref pVal, out BubbleEvent); break;
                    case "51115": prdDisc.MenuEvent_PrdDisc(ref oSBOCompany, ref oSBOApplication, ref pVal, out BubbleEvent); break;
                    case "32323": fixDisc.MenuEvent_FixDisc(ref oSBOCompany, ref oSBOApplication, ref pVal, out BubbleEvent); break;
                    
                    // Add
                    case "1282": MenuEventHandlerAdd(ref oSBOCompany, ref oSBOApplication, ref pVal, out BubbleEvent); break;
                    // next record
                    case "1288": MenuEventHandlerNextPrev(ref oSBOCompany, ref oSBOApplication, ref pVal, out BubbleEvent); break;
                    // previous record
                    case "1289": MenuEventHandlerNextPrev(ref oSBOCompany, ref oSBOApplication, ref pVal, out BubbleEvent); break;
                    // first data record
                    case "1290": MenuEventHandlerNextPrev(ref oSBOCompany, ref oSBOApplication, ref pVal, out BubbleEvent); break;
                    // last data record
                    case "1291": MenuEventHandlerNextPrev(ref oSBOCompany, ref oSBOApplication, ref pVal, out BubbleEvent); break;

                    case "PrdDiscAdd": prdDisc.MenuEvent_PrdDiscAdd(ref oSBOApplication, ref pVal, ref BubbleEvent); break;
                    case "PrdDiscDel": prdDisc.MenuEvent_PrdDiscDel(ref oSBOApplication, ref pVal, ref BubbleEvent); break;

                    case "FixDiscAdd": fixDisc.MenuEvent_FixDiscAdd(ref oSBOApplication, ref pVal, ref BubbleEvent); break;
                    case "FixDiscDel": fixDisc.MenuEvent_FixDiscDel(ref oSBOApplication, ref pVal, ref BubbleEvent); break;
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Handle Right Click Event
        /// </summary>
        public void HandleRightClickEvent(ref ContextMenuInfo eventInfo, out bool bubbleEvent)
        {
            Form oForm = oSBOApplication.Forms.ActiveForm;

            try
            {
                bubbleEvent = true;
                OneTimeDiscount oneTimeDisc = new OneTimeDiscount();
                PeriodicDiscount prdDisc = new PeriodicDiscount();
                FixDiscount fixDisc = new FixDiscount();

                switch (oForm.TypeEx)
                {
                    case "PRDDISC": prdDisc.RightClickEvent_PrdDisc(ref oSBOApplication, ref eventInfo, ref bubbleEvent); break;
                    case "FIXDISC": fixDisc.RightClickEvent_FixDisc(ref oSBOApplication, ref eventInfo, ref bubbleEvent); break;
                }
            }
            catch (Exception ex)
            {
                bubbleEvent = false;
                Utils.releaseObject(oForm);
                oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Handle Item Event
        /// </summary>
        public void HandleItemEvent(string FormUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            OneTimeDiscount cashDisc = new OneTimeDiscount();
            PeriodicDiscount prdDisc = new PeriodicDiscount();
            FixDiscount fixDisc = new FixDiscount();
            SalesOrder salesOrder = new SalesOrder();

            try
            {
                if (pVal.EventType != BoEventTypes.et_FORM_UNLOAD)
                {
                    if (pVal.FormTypeEx == "CASHDISC")
                        cashDisc.ItemEvent_CashDisc(ref oSBOCompany, ref oSBOApplication, FormUID, ref pVal, ref bubbleEvent);
                    else if (pVal.FormTypeEx == "PRDDISC")
                        prdDisc.ItemEvent_PrdDisc(ref oSBOCompany, ref oSBOApplication, FormUID, ref pVal, ref bubbleEvent);
                    else if (pVal.FormTypeEx == "FIXDISC")
                        fixDisc.ItemEvent_PrdDisc(ref oSBOCompany, ref oSBOApplication, FormUID, ref pVal, ref bubbleEvent);
                    else if (pVal.FormType == 139)
                        salesOrder.ItemEvents_SalesOrder(ref oSBOCompany, ref oSBOApplication, FormUID, ref pVal, ref bubbleEvent);
                }
            }
            catch (Exception ex)
            {
                bubbleEvent = false;
                oSBOApplication.MessageBox(ex.Message);
            }
        }

        /// <summary>
        /// Event when click Add Menu (CTRL+A)
        /// </summary>
        private void MenuEventHandlerAdd(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplicaton
                                        , ref SAPbouiCOM.MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (pVal.BeforeAction == false)
            {
                Form oForm = oSBOApplication.Forms.ActiveForm;

                PeriodicDiscount prdDisc = new PeriodicDiscount();
                FixDiscount fixDisc = new FixDiscount();

                switch (oForm.TypeEx)
                {
                    case "PRDDISC": prdDisc.Template_Add_PrdDisc(ref oSBOCompany, ref oSBOApplication, ref oForm); break;
                    case "FIXDISC": fixDisc.Template_Add_FixDisc(ref oSBOCompany, ref oSBOApplication, ref oForm); break;
                }

                Utils.releaseObject(oForm);
            }
        }

        private void MenuEventHandlerNextPrev(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplicaton
                                        , ref SAPbouiCOM.MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (pVal.BeforeAction == false)
            {
                Form oForm = oSBOApplication.Forms.ActiveForm;
                PeriodicDiscount prdDisc = new PeriodicDiscount();

                switch (oForm.TypeEx)
                {
                    case "PRDDISC": prdDisc.NextPrev_PrdDisc(ref oSBOCompany, ref oSBOApplication, ref oForm); break;
                }
            }
        }
    }
}
