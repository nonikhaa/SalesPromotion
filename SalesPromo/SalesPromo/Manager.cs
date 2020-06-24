using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;
using System.IO;

namespace SalesPromo
{
    public class Manager
    {
        public const string addonName = "SBO";

        private SAPbouiCOM.Application oSBOApplication = null;
        private SAPbobsCOM.Company oSBOCompany = null;

        public Manager()
        {
            StartUp();
        }

        private void StartUp()
        {
            try
            {
                SetupApplication();
                NumericSeparators();
                CatchingEvents();
                CreateUDT();
                CreateUDF();
                CreateUDO();
                CreateMenu();
                CreateFolder();
                CreateSP();
                CreateFMS();

                oSBOApplication.StatusBar.SetText(addonName + " Add-On Sales Promotion Connected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                if (oSBOApplication != null)
                    oSBOApplication.MessageBox(ex.Message);
                else
                    MessageBox.Show(ex.Message);

                System.Windows.Forms.Application.Exit();
            }
        }

        /// <summary>
        /// Connect to SAP
        /// </summary>
        private void SetupApplication()
        {
            SAPbouiCOM.SboGuiApi oSboGuiApi = null;
            string sConnectionString = null;

            oSboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            //sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

            oSboGuiApi.Connect(sConnectionString);

            oSBOApplication = oSboGuiApi.GetApplication();
            oSBOCompany = oSBOApplication.Company.GetDICompany();

            if (oSBOCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                GeneralVariables.SQLHandler = new HANAQueries();
            else
                GeneralVariables.SQLHandler = new SQLQueries();
        }

        /// <summary>
        /// Decimal Separator
        /// </summary>
        private void NumericSeparators()
        {
            SAPbobsCOM.Recordset oRec;

            oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec.DoQuery(GeneralVariables.SQLHandler.SeparatorSQL());

            GeneralVariables.WinDecSep = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            GeneralVariables.SBODecSep = oRec.Fields.Item("DecSep").Value.ToString();
            GeneralVariables.SBOThousSep = oRec.Fields.Item("ThousSep").Value.ToString();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            oRec = null;
            GC.Collect();
        }

        /// <summary>
        /// All Event
        /// </summary>
        public void CatchingEvents()
        {
            // events handled by SBO_Application_AppEvent 
            oSBOApplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBOApplication_AppEvent);

            // events handled by SBO_Application_MenuEvent 
            oSBOApplication.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBOApplication_MenuEvent);

            // events handled by SBO_Application_ItemEvent
            oSBOApplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBOApplication_ItemEvent);

            //// events handled by SBO_Application_ProgressBarEvent
            //oSBOApplication.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBOApplication_FormDataEvent);

            // events handled by SBO_Application_StatusBarEvent
            oSBOApplication.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBOApplication_RightClickEvent);
        }

        /// <summary>
        /// Create UDT
        /// </summary>
        private void CreateUDT()
        {
            Utils.CreateUDT(oSBOCompany, "SOL_CASHDISC", "One Time Discount", BoUTBTableType.bott_MasterData);

            Utils.CreateUDT(oSBOCompany, "SOL_PRDDISC_H", "Periodic Discount", BoUTBTableType.bott_MasterData);
            Utils.CreateUDT(oSBOCompany, "SOL_PRDDISC_D1", "Periodic Discount - Customer", BoUTBTableType.bott_MasterDataLines);
            Utils.CreateUDT(oSBOCompany, "SOL_PRDDISC_D2", "Periodic Discount - Discount %", BoUTBTableType.bott_MasterDataLines);
            Utils.CreateUDT(oSBOCompany, "SOL_PRDDISC_D3", "Periodic Discount - Buy1 Get1", BoUTBTableType.bott_MasterDataLines);

            Utils.CreateUDT(oSBOCompany, "SOL_FIXDISC_H", "Fix Discount", BoUTBTableType.bott_MasterData);
            Utils.CreateUDT(oSBOCompany, "SOL_FIXDISC_D1", "Fix Discount - Brand", BoUTBTableType.bott_MasterDataLines);
            Utils.CreateUDT(oSBOCompany, "SOL_FIXDISC_D2", "Fix Discount - Item", BoUTBTableType.bott_MasterDataLines);
        }

        /// <summary>
        /// Create UDF
        /// </summary>
        private void CreateUDF()
        {
            #region Cash Discount
            Utils.CreateUDF(oSBOCompany, "@SOL_CASHDISC", "SOL_CARDCODE", "Customer Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15);
            Utils.CreateUDF(oSBOCompany, "@SOL_CASHDISC", "SOL_CURRENCY", "Currency", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 5);
            Utils.CreateUDF(oSBOCompany, "@SOL_CASHDISC", "SOL_DPP", "Minimal DPP", BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 15);
            Utils.CreateUDF(oSBOCompany, "@SOL_CASHDISC", "SOL_CASHDISC", "Discount (%)", BoFieldTypes.db_Float, BoFldSubTypes.st_Percentage, 15);
            #endregion

            #region Periodic Discount
            // Header
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_H", "SOL_AREA", "Area", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_H", "SOL_CSTGRPCODE", "Customer Group Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 6);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_H", "SOL_CSTGRPNAME", "Customer Group Name", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_H", "SOL_STARTDATE", "Start Date", BoFieldTypes.db_Date, BoFldSubTypes.st_None, 10);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_H", "SOL_ENDDATE", "End Date", BoFieldTypes.db_Date, BoFldSubTypes.st_None, 10);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_H", "SOL_DISCTYPE", "Discount Type", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 5);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_H", "SOL_CUSTTYPE", "Customer Type", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15);

            // Detail 1
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D1", "SOL_CARDCODE", "Customer Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D1", "SOL_CARDNAME", "Customer Card Name", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100);

            // Detail 2
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D2", "SOL_ITEMCODE", "Item Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D2", "SOL_ITEMNAME", "Item Name", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D2", "SOL_MINQTY", "Min. Qty", BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, 18);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D2", "SOL_ITEMDISC", "Discount (%)", BoFieldTypes.db_Float, BoFldSubTypes.st_Percentage, 18);

            // Detail 3
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D3", "SOL_ITEMCODE", "Item Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D3", "SOL_ITEMNAME", "Item Name", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D3", "SOL_MINQTY", "Min. Qty", BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, 18);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D3", "SOL_ITMCD_FREE", "Item Code Bonus", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D3", "SOL_ITMNM_FREE", "Item Name Bonus", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D3", "SOL_QTY_FREE", "Qty. Bonus", BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, 18);
            Utils.CreateUDF(oSBOCompany, "@SOL_PRDDISC_D3", "SOL_KELIPATAN", "Kelipatan", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 5, null, "N");
            #endregion

            #region Fix Discount
            // Header
            Utils.CreateUDF(oSBOCompany, "@SOL_FIXDISC_H", "SOL_CARDCODE", "Customer Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15);
            Utils.CreateUDF(oSBOCompany, "@SOL_FIXDISC_H", "SOL_CARDNAME", "Customer Name", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100);
            Utils.CreateUDF(oSBOCompany, "@SOL_FIXDISC_H", "SOL_AREA", "Area", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50);

            // Detail 1
            Utils.CreateUDF(oSBOCompany, "@SOL_FIXDISC_D1", "SOL_BRANDCODE", "Brand Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2);
            Utils.CreateUDF(oSBOCompany, "@SOL_FIXDISC_D1", "SOL_BRANDNAME", "Brand Name", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50);
            Utils.CreateUDF(oSBOCompany, "@SOL_FIXDISC_D1", "SOL_BRANDDISC", "Discount (%)", BoFieldTypes.db_Float, BoFldSubTypes.st_Percentage, 18);

            // Detail 2
            Utils.CreateUDF(oSBOCompany, "@SOL_FIXDISC_D2", "SOL_ITEMCODE", "Item Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50);
            Utils.CreateUDF(oSBOCompany, "@SOL_FIXDISC_D2", "SOL_ITEMNAME", "Item Name", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100);
            Utils.CreateUDF(oSBOCompany, "@SOL_FIXDISC_D2", "SOL_ITEMDISC", "Discount (%)", BoFieldTypes.db_Float, BoFldSubTypes.st_Percentage, 18);
            #endregion

            #region Sales Order
            Utils.CreateUDF(oSBOCompany, "ORDR", "SOL_CASHDISC", "One Time Discount (%)", BoFieldTypes.db_Float, BoFldSubTypes.st_Percentage, 15);
            Utils.CreateUDF(oSBOCompany, "ORDR", "SOL_APPLDISC", "Apply Discount", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1);
            Utils.CreateUDF(oSBOCompany, "ORDR", "SOL_MEMO", "No. Memo", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30);

            Utils.CreateUDF(oSBOCompany, "RDR1", "SOL_ADDSC", "Add-On Disc (%)", BoFieldTypes.db_Float, BoFldSubTypes.st_Percentage, 15);
            Utils.CreateUDF(oSBOCompany, "RDR1", "SOL_FD", "Fix Disc (%)", BoFieldTypes.db_Float, BoFldSubTypes.st_Percentage, 15);
            Utils.CreateUDF(oSBOCompany, "RDR1", "SOL_FDCD", "Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 14);
            Utils.CreateUDF(oSBOCompany, "RDR1", "SOL_PD", "Periodic Disc (%)", BoFieldTypes.db_Float, BoFldSubTypes.st_Percentage, 15);
            Utils.CreateUDF(oSBOCompany, "RDR1", "SOL_PDCD", "Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 9);
            Utils.CreateUDF(oSBOCompany, "RDR1", "SOL_FLGBNS", "Bonus Item", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2);
            #endregion
        }

        #region Create UDO
        /// <summary>
        /// Create UDO
        /// </summary>
        private void CreateUDO()
        {
            CreateUDO_CashDisc();
            CreateUDO_PrdDisc();
            CreateUDO_FixDisc();
        }

        /// <summary>
        /// UDO - One Time Discount 
        /// </summary>
        private void CreateUDO_CashDisc()
        {
            string[] FormColumnAlias = { "Code", "U_SOL_CARDCODE", "U_SOL_CURRENCY", "U_SOL_DPP", "U_SOL_CASHDISC" };
            string[] FormColumnDescription = { "No. Memo", "Customer Code", "Currency", "Minimal DPP", "Discount(%)" };
            SAPbobsCOM.BoYesNoEnum[] FormColumnsEditable = { BoYesNoEnum.tYES, BoYesNoEnum.tYES, BoYesNoEnum.tYES,
                                                             BoYesNoEnum.tYES, BoYesNoEnum.tYES};

            Utils.CreateUDOTemplate(oSBOCompany, "CASHDISC", "One Time Discount", BoUDOObjType.boud_MasterData, "SOL_CASHDISC"
                               , BoYesNoEnum.tYES, BoYesNoEnum.tYES, BoYesNoEnum.tNO, BoYesNoEnum.tNO
                               , BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO
                               , BoYesNoEnum.tNO, BoYesNoEnum.tYES, "ASOL_CASHDISC", null, null, 0, 0
                               , FormColumnAlias, FormColumnDescription, FormColumnAlias, FormColumnDescription
                               , FormColumnsEditable, null);
        }

        /// <summary>
        /// UDO -  Periodic Discount
        /// </summary>
        private void CreateUDO_PrdDisc()
        {
            Utils.UDOChild child1 = new Utils.UDOChild("SOL_PRDDISC_D1");
            Utils.UDOChild child2 = new Utils.UDOChild("SOL_PRDDISC_D2");
            Utils.UDOChild child3 = new Utils.UDOChild("SOL_PRDDISC_D3");
            Utils.UDOChild[] childs = { child1, child2, child3 };

            string[] FormColumnAlias = {"Code", "U_SOL_AREA", "U_SOL_CSTGRPCODE", "U_SOL_CSTGRPNAME"
                                        , "U_SOL_STARTDATE", "U_SOL_ENDDATE", "U_SOL_DISCTYPE"};

            string[] FormColumnDescription = {"Code", "Area", "Customer Group Code", "Customer Group Name"
                                               , "Start Date", "End Date", "Discount Type"};

            SAPbobsCOM.BoYesNoEnum[] FormColumnEditable = { BoYesNoEnum.tNO, BoYesNoEnum.tYES, BoYesNoEnum.tYES
                                                            , BoYesNoEnum.tYES, BoYesNoEnum.tYES, BoYesNoEnum.tYES
                                                            , BoYesNoEnum.tYES};

            Utils.CreateUDOTemplate(oSBOCompany, "PRDDISC", "Periodic Discount Master Data", BoUDOObjType.boud_MasterData, "SOL_PRDDISC_H"
                              , BoYesNoEnum.tYES, BoYesNoEnum.tYES, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO
                              , BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tYES
                              , "ASOL_PRDDISC_H", null, null, 0, 0, FormColumnAlias, FormColumnDescription, FormColumnAlias
                              , FormColumnDescription, FormColumnEditable, childs);
        }

        /// <summary>
        /// UDO - Fix Discount
        /// </summary>
        private void CreateUDO_FixDisc()
        {
            Utils.UDOChild child1 = new Utils.UDOChild("SOL_FIXDISC_D1");
            Utils.UDOChild child2 = new Utils.UDOChild("SOL_FIXDISC_D2");
            Utils.UDOChild[] childs = { child1, child2 };

            string[] FormColumnAlias = { "Code", "U_SOL_AREA", "U_SOL_CARDCODE", "U_SOL_CARDNAME" };
            string[] FormColumnDescription = { "Code", "Area", "Customer Code", "Customer Name" };
            SAPbobsCOM.BoYesNoEnum[] FormColumnEditable = { BoYesNoEnum.tNO, BoYesNoEnum.tYES, BoYesNoEnum.tYES
                                                            , BoYesNoEnum.tYES};

            Utils.CreateUDOTemplate(oSBOCompany, "FIXDISC", "Fix Discount Master Data", BoUDOObjType.boud_MasterData, "SOL_FIXDISC_H"
                              , BoYesNoEnum.tYES, BoYesNoEnum.tYES, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO
                              , BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tYES
                              , "ASOL_FIXDISC_H", null, null, 0, 0, FormColumnAlias, FormColumnDescription, FormColumnAlias
                              , FormColumnDescription, FormColumnEditable, childs);
        }

        #endregion

        /// <summary>
        /// Create Menu
        /// </summary>
        private void CreateMenu()
        {
            Utils.CreateMenuEntryTemplate(ref oSBOApplication, ref oSBOCompany, "78787", "Sales Promotion Add-On", BoMenuType.mt_POPUP);
            Utils.CreateMenuEntryTemplate(ref oSBOApplication, ref oSBOCompany, "96969", "One Time Discount Setting", BoMenuType.mt_STRING, 0, "78787");
            Utils.CreateMenuEntryTemplate(ref oSBOApplication, ref oSBOCompany, "32323", "Fix Discount Master Data", BoMenuType.mt_STRING, 1, "78787");
            Utils.CreateMenuEntryTemplate(ref oSBOApplication, ref oSBOCompany, "51115", "Periodic Discount", BoMenuType.mt_STRING, 2, "78787");
        }

        /// <summary>
        /// Create folder for put FMS or SP
        /// </summary>
        private void CreateFolder()
        {
            string dir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string dirPathSP = dir + @"\SP\";
            string dirPathFMS = dir + @"\FMS\";

            if (!Directory.Exists(dirPathSP))
            {
                Directory.CreateDirectory(dirPathSP);
            }

            if (!Directory.Exists(dirPathFMS))
            {
                Directory.CreateDirectory(dirPathFMS);
            }
        }

        #region Create FMS
        /// <summary>
        /// Formated Search (FMS)
        /// </summary>
        private void CreateFMS()
        {
            CreateFMS_PrdDisc();
            CreateFMS_FixDisc();
        }

        /// <summary>
        /// FMS - Periodic Discount
        /// </summary>
        private void CreateFMS_PrdDisc()
        {
            Utils.CreateQueryCategory(oSBOCompany, "ADDON - Periodic Discount");
            string queryCategory = "ADDON - Periodic Discount";
            string formId = "PRDDISC";

            if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
            {
                Utils.CreateQueryFromText(oSBOCompany, oSBOApplication, queryCategory, "SOL - PRDDISC - Area", "HANA - SOL - PRDDISC - Area.sql");
            }
            else
            {
                Utils.CreateQueryFromText(oSBOCompany, oSBOApplication, queryCategory, "SOL - PRDDISC - Area", "SQL - SOL - PRDDISC - Area.sql");
            }

            // Create query from file
            Utils.CreateQueryFromText(oSBOCompany, oSBOApplication, queryCategory, "SOL - PRDDISC - Customer Code", "SOL - PRDDISC - Customer Code.sql");
            Utils.CreateQueryFromText(oSBOCompany, oSBOApplication, queryCategory, "SOL - PRDDISC - Item Code", "SOL - PRDDISC - Item Code.sql");
            Utils.CreateQueryFromText(oSBOCompany, oSBOApplication, queryCategory, "SOL - PRDDISC - Item Name", "SOL - PRDDISC - Item Name.sql");

            // Create FMS
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - PRDDISC - Area", formId, "tArea", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO);
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - PRDDISC - Customer Code", formId, "mt_1", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, "cCsCd");
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - PRDDISC - Item Code", formId, "mt_2", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, "cItmCd");
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - PRDDISC - Item Name", formId, "mt_2", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, "cItmNm");
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - PRDDISC - Item Code", formId, "mt_3", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, "cItmC");
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - PRDDISC - Item Code", formId, "mt_3", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, "cItmCB");
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - PRDDISC - Item Name", formId, "mt_3", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, "cItmN");
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - PRDDISC - Item Name", formId, "mt_3", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, "cItmNB");
        }

        /// <summary>
        /// FMS - Fix Discount
        /// </summary>
        private void CreateFMS_FixDisc()
        {
            Utils.CreateQueryCategory(oSBOCompany, "ADDON - Fix Discount");
            string queryCategory = "ADDON - Fix Discount";
            string formId = "FIXDISC";

            if(oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
            {
                Utils.CreateQueryFromText(oSBOCompany, oSBOApplication, queryCategory, "SOL - FIXDISC - Area", "HANA - SOL - FIXDISC - Area.sql");
            }
            else
            {
                Utils.CreateQueryFromText(oSBOCompany, oSBOApplication, queryCategory, "SOL - FIXDISC - Area", "SQL - SOL - FIXDISC - Area.sql");
            }
            
            Utils.CreateQueryFromText(oSBOCompany, oSBOApplication, queryCategory, "SOL - FIXDISC - Brand Code", "SOL - FIXDISC - Brand Code.sql");
            Utils.CreateQueryFromText(oSBOCompany, oSBOApplication, queryCategory, "SOL - FIXDISC - Item Code", "SOL - FIXDISC - Item Code.sql");

            // Create FMS
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - FIXDISC - Area", formId, "tArea", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO);
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - FIXDISC - Brand Code", formId, "mt_1", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, "cBrCd");
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - FIXDISC - Item Code", formId, "mt_2", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, "cItmCd");
        }

        #endregion

        #region Create SP
        /// <summary>
        /// Create Stored Procedure (SP) - load from file
        /// </summary>
        private void CreateSP()
        {
            CreateSP_CashDisc();
            CreateSP_PrdDisc();
            CreateSP_FixDisc();
            CreateSP_CalculateSO();

            // Call/Exec ALL SP
            CreateSPValidation();
        }

        /// <summary>
        /// SP Periodic discount
        /// </summary>
        private void CreateSP_PrdDisc()
        {
            try
            {
                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                {
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SBO_SP_TRANSACTIONNOTIFICATION_PRDDISC.sql", "SBO_SP_TRANSACTIONNOTIFICATION_PRDDISC");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_ADDON_PD_CODE.sql", "SOL_SP_ADDON_PD_CODE");
                }
                else
                {
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "SQL - SBO_SP_TRANSACTIONNOTIFICATION_PRDDISC.sql", "SBO_SP_TRANSACTIONNOTIFICATION_PRDDISC");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "SQL - SOL_SP_ADDON_PD_CODE.sql", "SOL_SP_ADDON_PD_CODE");
                }
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
        }

        /// <summary>
        /// SP One Time Discount (Cash Discount)
        /// </summary>
        private void CreateSP_CashDisc()
        {
            try
            {
                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                {
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SBO_SP_TRANSACTIONNOTIFICATION_CASHDISC.sql", "SBO_SP_TRANSACTIONNOTIFICATION_CASHDISC");
                }
                else
                {
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "SQL - SBO_SP_TRANSACTIONNOTIFICATION_CASHDISC.sql", "SBO_SP_TRANSACTIONNOTIFICATION_CASHDISC");
                }
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
        }

        /// <summary>
        /// SP Fix Discount Master Data
        /// </summary>
        private void CreateSP_FixDisc()
        {
            try
            {
                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                {
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SBO_SP_TRANSACTIONNOTIFICATION_FIXDISC.sql", "SBO_SP_TRANSACTIONNOTIFICATION_FIXDISC");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_ADDON_FD_CODE.sql", "SOL_SP_ADDON_FD_CODE");
                }
                else
                {
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "SQL - SBO_SP_TRANSACTIONNOTIFICATION_FIXDISC.sql", "SBO_SP_TRANSACTIONNOTIFICATION_FIXDISC");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "SQL - SOL_SP_ADDON_FD_CODE.sql", "SOL_SP_ADDON_FD_CODE");
                }
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
        }

        /// <summary>
        /// SP Calculation in Sales Order
        /// </summary>
        private void CreateSP_CalculateSO()
        {
            try
            {
                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                {
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_ADDON_GET_PRDDISC.sql", "SOL_SP_ADDON_GET_PRDDISC");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_ADDON_GET_CASHDISC.sql", "SOL_SP_ADDON_GET_CASHDISC");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_ADDON_GET_FIXDISC.sql", "SOL_SP_ADDON_GET_FIXDISC");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SBO_SP_TRANSACTIONNOTIFICATION_ADDONSO.sql", "SBO_SP_TRANSACTIONNOTIFICATION_ADDONSO");
                }
                else
                {
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "SQL - SOL_SP_ADDON_GET_FIXDISC.sql", "SOL_SP_ADDON_GET_FIXDISC");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "SQL - SOL_SP_ADDON_GET_CASHDISC.sql", "SOL_SP_ADDON_GET_CASHDISC");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "SQL - SOL_SP_ADDON_GET_PRDDISC.sql", "SOL_SP_ADDON_GET_PRDDISC");
                }
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
        }

        private void CreateSPValidation()
        {
            string qCustomSPValidation = string.Empty;
            string qSPValidation = string.Empty;
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                string dir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string dirPathSP = dir + @"\SP\";

                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                {
                    #region HANA
                    string fileName = Path.Combine(dirPathSP, Path.GetFileName("HANA - SBO_SP_TransactionNotification.sql"));

                    if (File.Exists(fileName))
                    {
                        using (StreamReader sr = new StreamReader(fileName))
                        {
                            qCustomSPValidation = sr.ReadToEnd();
                        }
                    }

                    oRec.DoQuery("SELECT DEFINITION FROM SYS.PROCEDURES WHERE SCHEMA_NAME = '" + oSBOCompany.CompanyDB + "' AND UPPER(PROCEDURE_NAME) = 'SBO_SP_TRANSACTIONNOTIFICATION'");
                    if (oRec.RecordCount > 0)
                    {
                        qSPValidation = oRec.Fields.Item("DEFINITION").Value;
                        qCustomSPValidation = qCustomSPValidation.Trim();
                        qSPValidation = qSPValidation.Trim();
                        if (qSPValidation.ToLower().Contains("-- Add-On Sales Promotion (SOLTIUS)".ToLower()) == false)
                        {
                            string a = "-- Select the return values";
                            string b = "select :error, :error_message FROM dummy";

                            qSPValidation = qSPValidation.Replace("CREATE PROCEDURE", "ALTER PROCEDURE");

                            qCustomSPValidation = qCustomSPValidation + System.Environment.NewLine;
                            qCustomSPValidation = qCustomSPValidation + System.Environment.NewLine + a;
                            qCustomSPValidation = qCustomSPValidation + System.Environment.NewLine + b;

                            if (qSPValidation.ToLower().Contains(a.ToLower()))
                            {
                                qSPValidation = qSPValidation.Substring(0, qSPValidation.ToLower().LastIndexOf(a.ToLower())) +
                                qSPValidation.Substring(qSPValidation.ToLower().LastIndexOf(a.ToLower()) + a.Length);
                            }

                            qSPValidation = qSPValidation.Substring(0, qSPValidation.ToLower().LastIndexOf(b.ToLower()))
                                            + qCustomSPValidation
                                            + qSPValidation.Substring(qSPValidation.ToLower().LastIndexOf(b.ToLower()) + b.Length);
                            oRec.DoQuery(qSPValidation);
                        }
                    }
                    #endregion
                }
                else
                {
                    #region SQL
                    string fileName = Path.Combine(dirPathSP, Path.GetFileName("SQL - SBO_SP_TransactionNotification.sql"));

                    if (File.Exists(fileName))
                    {
                        using (StreamReader sr = new StreamReader(fileName))
                        {
                            qCustomSPValidation = sr.ReadToEnd();
                        }
                    }

                    oRec.DoQuery("SELECT definition FROM sys.sql_modules WHERE objectproperty(object_id, 'IsProcedure') = 1 AND UPPER(OBJECT_NAME(object_id)) = 'SBO_SP_TRANSACTIONNOTIFICATION'");
                    if (oRec.RecordCount > 0)
                    {
                        qSPValidation = oRec.Fields.Item("definition").Value;
                        qCustomSPValidation = qCustomSPValidation.Trim();
                        qSPValidation = qSPValidation.Trim();

                        if (qSPValidation.ToLower().Contains("-- Add-On Sales Promotion (SOLTIUS)".ToLower()) == false)
                        {
                            string a = "-- Select the return values";
                            string b = "select @error, @error_message";

                            qSPValidation = qSPValidation.ToLower().Replace("create procedure", "alter procedure");
                            qSPValidation = qSPValidation.ToLower().Replace("create proc", "alter proc");

                            qCustomSPValidation = qCustomSPValidation + System.Environment.NewLine;
                            qCustomSPValidation = qCustomSPValidation + System.Environment.NewLine + a;
                            qCustomSPValidation = qCustomSPValidation + System.Environment.NewLine + b;

                            if (qSPValidation.ToLower().Contains(a.ToLower()))
                            {
                                qSPValidation = qSPValidation.Substring(0, qSPValidation.ToLower().LastIndexOf(a.ToLower())) +
                                qSPValidation.Substring(qSPValidation.ToLower().LastIndexOf(a.ToLower()) + a.Length);
                            }

                            qSPValidation = qSPValidation.Substring(0, qSPValidation.ToLower().LastIndexOf(b.ToLower()))
                                            + qCustomSPValidation
                                            + qSPValidation.Substring(qSPValidation.ToLower().LastIndexOf(b.ToLower()) + b.Length);
                            oRec.DoQuery(qSPValidation);
                        }
                    }
                    #endregion
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
        }

        #endregion

        #region SBO Event Handler
        private void SBOApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
            oSBOEventHandler.HandleAppEvent(EventType);
        }

        private void SBOApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
            oSBOEventHandler.HandleMenuEvent(ref pVal, out BubbleEvent);
        }

        private void SBOApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
            oSBOEventHandler.HandleItemEvent(FormUID, ref pVal, out BubbleEvent);
        }

        //private void SBOApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        //{
        //    SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
        //    oSBOEventHandler.HandleFormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
        //}

        private void SBOApplication_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
            oSBOEventHandler.HandleRightClickEvent(ref eventInfo, out BubbleEvent);
        }
        #endregion
    }
}
