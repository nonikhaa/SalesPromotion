using SAPbouiCOM;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Xml;
using System.IO;

namespace SalesPromo
{
    public static class Utils
    {
        public static void CreateUDF(SAPbobsCOM.Company oSBOCompany, string TableName, string FieldName, string FieldDescription, SAPbobsCOM.BoFieldTypes Type)
        {
            CreateUDF(oSBOCompany, TableName, FieldName, FieldDescription, Type, SAPbobsCOM.BoFldSubTypes.st_None);
        }

        public static void CreateUDF(SAPbobsCOM.Company oSBOCompany, string TableName, string FieldName, string FieldDescription, SAPbobsCOM.BoFieldTypes Type, SAPbobsCOM.BoFldSubTypes SubType)
        {
            CreateUDF(oSBOCompany, TableName, FieldName, FieldDescription, Type, SubType, 0);
        }

        public static void CreateUDF(SAPbobsCOM.Company oSBOCompany, string TableName, string FieldName, string FieldDescription, SAPbobsCOM.BoFieldTypes Type, int EditSize)
        {
            CreateUDF(oSBOCompany, TableName, FieldName, FieldDescription, Type, SAPbobsCOM.BoFldSubTypes.st_None, EditSize);
        }

        public static void CreateUDF(SAPbobsCOM.Company oSBOCompany, string TableName, string FieldName, string FieldDescription, SAPbobsCOM.BoFieldTypes Type, SAPbobsCOM.BoFldSubTypes SubType, int EditSize)
        {
            CreateUDF(oSBOCompany, TableName, FieldName, FieldDescription, Type, SubType, EditSize, new List<string[]>(), "");
        }

        public static void CreateUDF(SAPbobsCOM.Company oSBOCompany, string TableName, string FieldName, string FieldDescription, SAPbobsCOM.BoFieldTypes Type, SAPbobsCOM.BoFldSubTypes SubType, int EditSize, List<string[]> ValidValues, string DefaultValue)
        {
            CreateUDF(oSBOCompany, TableName, FieldName, FieldDescription, Type, SubType, EditSize, ValidValues, null, DefaultValue, SAPbobsCOM.BoYesNoEnum.tNO);
        }

        public static void CreateUDF(SAPbobsCOM.Company oSBOCompany, string TableName, string FieldName, string FieldDescription, SAPbobsCOM.BoFieldTypes Type, SAPbobsCOM.BoFldSubTypes SubType, int EditSize, List<string[]> ValidValues, string LinkedTable, string DefaultValue, SAPbobsCOM.BoYesNoEnum Mandatory)
        {
            SAPbobsCOM.UserFieldsMD oUFields;
            SAPbobsCOM.Recordset oRec;
            int lRetCode = 0;
            int lErrCode = 0;
            string sErrMsg = "";
            bool bIsContinue = true;

            oRec = (SAPbobsCOM.Recordset)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec.DoQuery(GeneralVariables.SQLHandler.CheckUDFSQL(TableName, FieldName));

            if (oRec.RecordCount == 0)
                bIsContinue = true;
            else
                bIsContinue = false;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            oRec = null;
            GC.Collect();

            if (bIsContinue)
            {
                oUFields = (SAPbobsCOM.UserFieldsMD)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUFields.TableName = TableName;
                oUFields.Name = FieldName;
                oUFields.Description = FieldDescription;
                oUFields.Type = Type;
                oUFields.SubType = SubType;

                if (EditSize > 0)
                    oUFields.EditSize = EditSize;

                if (ValidValues != null)
                {
                    for (int i = 0; i <= ValidValues.Count - 1; i++)
                    {
                        string[] values = ValidValues.ElementAt(i);

                        oUFields.ValidValues.Value = values[0];
                        oUFields.ValidValues.Description = values[1];
                        oUFields.ValidValues.Add();
                    }
                }

                oUFields.LinkedTable = LinkedTable;
                oUFields.DefaultValue = DefaultValue;
                oUFields.Mandatory = Mandatory;

                lRetCode = oUFields.Add();

                if (lRetCode != 0)
                {
                    lErrCode = oSBOCompany.GetLastErrorCode();
                    sErrMsg = oSBOCompany.GetLastErrorDescription();
                    throw new Exception(lErrCode.ToString() + " : " + sErrMsg);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields);
                oUFields = null;
                GC.Collect();
            }
        }

        //public static void CreateUDF(SAPbobsCOM.Company oSBOCompany, string TableName, string FieldName, string FieldDescription, SAPbobsCOM.BoFieldTypes Type, SAPbobsCOM.BoFldSubTypes SubType, int EditSize, string LinkedTable, string DefaultValue, SAPbobsCOM.BoYesNoEnum Mandatory)
        //{
        //    SAPbobsCOM.UserFieldsMD oUFields;
        //    SAPbobsCOM.Recordset oRec;
        //    int lRetCode = 0;
        //    int lErrCode = 0;
        //    string sErrMsg = "";
        //    bool bIsContinue = true;

        //    oRec = (SAPbobsCOM.Recordset)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    oRec.DoQuery(GeneralVariables.SQLHandler.CheckUDFSQL(TableName, FieldName));

        //    if (oRec.RecordCount == 0)
        //        bIsContinue = true;
        //    else
        //        bIsContinue = false;

        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
        //    oRec = null;
        //    GC.Collect();

        //    if (bIsContinue)
        //    {
        //        oUFields = (SAPbobsCOM.UserFieldsMD)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
        //        oUFields.TableName = TableName;
        //        oUFields.Name = FieldName;
        //        oUFields.Description = FieldDescription;
        //        oUFields.Type = Type;
        //        oUFields.SubType = SubType;

        //        if (EditSize > 0)
        //            oUFields.EditSize = EditSize;

        //        oUFields.LinkedTable = LinkedTable;
        //        oUFields.DefaultValue = DefaultValue;
        //        oUFields.Mandatory = Mandatory;

        //        lRetCode = oUFields.Add();

        //        if (lRetCode != 0)
        //        {
        //            lErrCode = oSBOCompany.GetLastErrorCode();
        //            sErrMsg = oSBOCompany.GetLastErrorDescription();
        //            throw new Exception(lErrCode.ToString() + " : " + sErrMsg);
        //        }

        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields);
        //        oUFields = null;
        //        GC.Collect();
        //    }
        //}

        public static void CreateUDT(SAPbobsCOM.Company oSBOCompany, string TableName, string TableDescription, SAPbobsCOM.BoUTBTableType TableType)
        {
            SAPbobsCOM.UserTablesMD oUTables;
            int lRetCode = 0;
            int lErrCode = 0;
            string sErrMsg = "";

            oUTables = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

            if (!oUTables.GetByKey(TableName))
            {
                oUTables.TableName = TableName;
                oUTables.TableDescription = TableDescription;
                oUTables.TableType = TableType;
                lRetCode = oUTables.Add();
            }

            if (lRetCode != 0)
            {
                lErrCode = oSBOCompany.GetLastErrorCode();
                sErrMsg = oSBOCompany.GetLastErrorDescription();
                throw new Exception(lErrCode.ToString() + " : " + sErrMsg);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUTables);
            oUTables = null;
            GC.Collect();
        }

        public static void CreateMenu(SAPbobsCOM.Company oSBOCompany, SAPbouiCOM.Application oSBOApplication, string ParentMenuUID, string MenuUID, SAPbouiCOM.BoMenuType MenuType, string MenuName, string MenuImage = null, int MenuPosition = 0, bool DeleteIfExists = false)
        {
            SAPbouiCOM.MenuItem oMenuItem;
            SAPbouiCOM.Menus oMenus;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            oCreationPackage = oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

            oMenuItem = oSBOApplication.Menus.Item(ParentMenuUID);
            oMenus = oMenuItem.SubMenus;

            string sPath = System.Windows.Forms.Application.StartupPath + @"\";

            bool MenuExists = false;

            for (int i = 0; i <= oMenus.Count - 1; i++)
            {
                if (oMenus.Exists(MenuUID))
                {
                    MenuExists = true;
                    break;
                }
            }

            if (MenuExists)
            {
                if (DeleteIfExists)
                {
                    oMenus.RemoveEx(MenuUID);

                    oCreationPackage.Type = MenuType;
                    oCreationPackage.UniqueID = MenuUID;
                    oCreationPackage.String = MenuName;
                    oCreationPackage.Image = sPath + MenuImage;
                    oCreationPackage.Position = MenuPosition;

                    oMenus.AddEx(oCreationPackage);
                }
            }
            else
            {
                oCreationPackage.Type = MenuType;
                oCreationPackage.UniqueID = MenuUID;
                oCreationPackage.String = MenuName;
                oCreationPackage.Image = sPath + MenuImage;
                oCreationPackage.Position = MenuPosition;

                oMenus.AddEx(oCreationPackage);
            }
        }

        public static void CreateUDO(SAPbobsCOM.Company oSBOCompany, string ObjectName, SAPbobsCOM.BoUDOObjType ObjectType
            , string TableName, SAPbobsCOM.BoYesNoEnum CanApprove, SAPbobsCOM.BoYesNoEnum CanArchive
            , SAPbobsCOM.BoYesNoEnum CanCancel, SAPbobsCOM.BoYesNoEnum CanClose, SAPbobsCOM.BoYesNoEnum CanCreateDefaultForm
            , SAPbobsCOM.BoYesNoEnum CanDelete, SAPbobsCOM.BoYesNoEnum CanFind, SAPbobsCOM.BoYesNoEnum CanLog
            , SAPbobsCOM.BoYesNoEnum CanYearTransfer, SAPbobsCOM.BoYesNoEnum ManageSeries, string[] FindColumns
            , string[] ChildTables)
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD;
            SAPbobsCOM.Recordset oRec;
            int lRetCode = 0;
            int lErrCode = 0;
            string sErrMsg = "";
            bool bIsContinue = true;

            oRec = (SAPbobsCOM.Recordset)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec.DoQuery(GeneralVariables.SQLHandler.CheckUDOSQL(ObjectName));

            if (oRec.RecordCount == 0)
                bIsContinue = true;
            else
                bIsContinue = false;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            oRec = null;
            GC.Collect();

            if (bIsContinue)
            {
                oUserObjectMD = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

                oUserObjectMD.Code = ObjectName;
                oUserObjectMD.Name = ObjectName;
                oUserObjectMD.ObjectType = ObjectType;
                oUserObjectMD.TableName = TableName;

                oUserObjectMD.CanApprove = CanApprove;
                oUserObjectMD.CanArchive = CanArchive;
                oUserObjectMD.CanCancel = CanCancel;
                oUserObjectMD.CanClose = CanClose;
                oUserObjectMD.CanCreateDefaultForm = CanCreateDefaultForm;
                oUserObjectMD.CanDelete = CanDelete;
                oUserObjectMD.CanFind = CanFind;
                oUserObjectMD.CanLog = CanLog;
                oUserObjectMD.CanYearTransfer = CanYearTransfer;

                oUserObjectMD.ManageSeries = ManageSeries;

                for (int i = 0; i <= FindColumns.Length - 1; i++)
                {
                    oUserObjectMD.FindColumns.ColumnAlias = FindColumns[i];
                    oUserObjectMD.FindColumns.Add();
                }

                for (int i = 0; i <= ChildTables.Length - 1; i++)
                {
                    oUserObjectMD.ChildTables.TableName = ChildTables[i];
                    oUserObjectMD.ChildTables.Add();
                }

                lRetCode = oUserObjectMD.Add();

                if (lRetCode != 0)
                {
                    lErrCode = oSBOCompany.GetLastErrorCode();
                    sErrMsg = oSBOCompany.GetLastErrorDescription();
                    throw new Exception(lErrCode.ToString() + " : " + sErrMsg);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;

                GC.Collect(); // Release the handle to the table
            }
        }

        public static void CreateUDO(SAPbobsCOM.Company oSBOCompany, string ObjectName, SAPbobsCOM.BoUDOObjType ObjectType,
            string TableName, SAPbobsCOM.BoYesNoEnum CanApprove, SAPbobsCOM.BoYesNoEnum CanArchive
            , SAPbobsCOM.BoYesNoEnum CanCancel, SAPbobsCOM.BoYesNoEnum CanClose, SAPbobsCOM.BoYesNoEnum CanCreateDefaultForm
            , SAPbobsCOM.BoYesNoEnum CanDelete, SAPbobsCOM.BoYesNoEnum CanFind, SAPbobsCOM.BoYesNoEnum CanLog
            , SAPbobsCOM.BoYesNoEnum CanYearTransfer, SAPbobsCOM.BoYesNoEnum ManageSeries, string FatherMenuID
            , string MenuCaption, string[] FindColumns, string[] ChildTables, string[,] FormColumns)
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD;
            SAPbobsCOM.Recordset oRec;
            int lRetCode = 0;
            int lErrCode = 0;
            string sErrMsg = "";
            bool bIsContinue = true;

            oRec = (SAPbobsCOM.Recordset)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec.DoQuery(GeneralVariables.SQLHandler.CheckUDOSQL(ObjectName));

            if (oRec.RecordCount == 0)
                bIsContinue = true;
            else
                bIsContinue = false;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            oRec = null;
            GC.Collect();

            if (bIsContinue)
            {
                oUserObjectMD = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

                oUserObjectMD.Code = ObjectName;
                oUserObjectMD.Name = ObjectName;
                oUserObjectMD.ObjectType = ObjectType;
                oUserObjectMD.TableName = TableName;

                oUserObjectMD.CanApprove = CanApprove;
                oUserObjectMD.CanArchive = CanArchive;
                oUserObjectMD.CanCancel = CanCancel;
                oUserObjectMD.CanClose = CanClose;

                oUserObjectMD.CanDelete = CanDelete;
                oUserObjectMD.CanFind = CanFind;
                oUserObjectMD.CanLog = CanLog;
                oUserObjectMD.CanYearTransfer = CanYearTransfer;

                oUserObjectMD.ManageSeries = ManageSeries;

                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.MenuItem = CanCreateDefaultForm;
                oUserObjectMD.FatherMenuID = int.Parse(FatherMenuID);
                oUserObjectMD.MenuCaption = MenuCaption;
                oUserObjectMD.Position = 0;
                oUserObjectMD.MenuUID = ObjectName;

                for (int i = 0; i <= FindColumns.Length - 1; i++)
                {
                    oUserObjectMD.FindColumns.ColumnAlias = FindColumns[i];
                    oUserObjectMD.FindColumns.Add();
                }

                for (int i = 0; i <= FormColumns.GetLength(0) - 1; i++)
                {
                    oUserObjectMD.FormColumns.FormColumnAlias = FormColumns[i, 0];
                    oUserObjectMD.FormColumns.FormColumnDescription = FormColumns[i, 1];
                    oUserObjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUserObjectMD.FormColumns.Add();
                }

                for (int i = 0; i <= ChildTables.Length - 1; i++)
                {
                    oUserObjectMD.ChildTables.TableName = ChildTables[i];
                    oUserObjectMD.ChildTables.Add();
                }

                lRetCode = oUserObjectMD.Add();

                if (lRetCode != 0)
                {
                    lErrCode = oSBOCompany.GetLastErrorCode();
                    sErrMsg = oSBOCompany.GetLastErrorDescription();
                    throw new Exception(lErrCode.ToString() + " : " + sErrMsg);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;

                GC.Collect(); // Release the handle to the table
            }
        }

        public static void CreateUDOTemplate(ref SAPbobsCOM.Company oCompany
                             , ref SAPbouiCOM.Application oSBOApplication
                             , string UDOCode
                             , string UDOName
                             , SAPbobsCOM.BoUDOObjType ObjectType
                             , string TableName
                             , SAPbobsCOM.BoYesNoEnum CanFind
                             , SAPbobsCOM.BoYesNoEnum CanDelete
                             , SAPbobsCOM.BoYesNoEnum CanCancel
                             , SAPbobsCOM.BoYesNoEnum CanClose
                             , SAPbobsCOM.BoYesNoEnum ManageSeries
                             , SAPbobsCOM.BoYesNoEnum CanYearTransfer
                             , SAPbobsCOM.BoYesNoEnum CanArchive
                             , SAPbobsCOM.BoYesNoEnum CanCreateDefaultForm
                             , SAPbobsCOM.BoYesNoEnum CanLog
                             , string LogTableName = null
                             , string MenuCaption = null
                             , int FatherMenuID = default(int)
                             , int Position = default(int)
                             , string MenuUID = null
                             , string[] FindColumnAlias = null
                             , string[] FindColumnDescription = null
                             , string[] FormColumnAlias = null
                             , string[] FormColumnDescription = null
                             , SAPbobsCOM.BoYesNoEnum[] FormEditable = null
                             , string ChildTableName = null
                             , string[] EnhancedFormColumnAlias = null
                             , string[] EnhancedFormColumnDescription = null
                             , SAPbobsCOM.BoYesNoEnum[] EnhancedFormColumnIsUsed = null
                             , SAPbobsCOM.BoYesNoEnum[] EnhancedFormEditable = null
                             , string ChildTableName2 = null
                             , string[] EnhancedFormColumnAlias2 = null
                             , string[] EnhancedFormColumnDescription2 = null
                             , SAPbobsCOM.BoYesNoEnum[] EnhancedFormColumnIsUsed2 = null
                             , SAPbobsCOM.BoYesNoEnum[] EnhancedFormEditable2 = null
                             )
        {
            SAPbobsCOM.UserObjectsMD oUDO;
            SAPbobsCOM.Recordset oRec;
            int ret, ErrCode;
            string ErrMsg = "";
            bool Result = true;


            try
            {
                // --------------------------------------------------------------------------------------------
                oRec = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery("SELECT \"Code\" FROM OUDO WHERE \"Code\" = '" + UDOCode + "'");

                if (oRec.RecordCount == 0)
                    Result = true;
                else
                    Result = false;

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
                oRec = null/* TODO Change to default(_) if this is not a reference type */;
                GC.Collect();

                if (Result)
                {
                    oUDO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                    oUDO.Code = UDOCode;
                    oUDO.Name = UDOName;
                    oUDO.ObjectType = ObjectType;
                    oUDO.TableName = TableName;
                    oUDO.CanFind = CanFind;
                    oUDO.CanDelete = CanDelete;
                    oUDO.CanCancel = CanCancel;
                    oUDO.CanClose = CanClose;
                    oUDO.ManageSeries = ManageSeries;
                    oUDO.CanYearTransfer = CanYearTransfer;
                    oUDO.CanArchive = CanArchive;
                    oUDO.CanCreateDefaultForm = CanCreateDefaultForm;
                    if (CanLog == SAPbobsCOM.BoYesNoEnum.tYES)
                    {
                        oUDO.CanLog = CanLog;
                        oUDO.LogTableName = LogTableName;
                    }

                    if (MenuCaption != null)
                    {
                        oUDO.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO;
                        oUDO.MenuItem = SAPbobsCOM.BoYesNoEnum.tYES;
                        oUDO.MenuCaption = MenuCaption;
                        oUDO.FatherMenuID = FatherMenuID;
                        oUDO.Position = Position;
                        oUDO.MenuUID = MenuUID;
                    }
                    else
                        oUDO.MenuItem = SAPbobsCOM.BoYesNoEnum.tNO;

                    if (FindColumnAlias != null)
                    {
                        for (int i = 0; i <= FindColumnAlias.Count() - 1; i++)
                        {
                            oUDO.FindColumns.ColumnAlias = FindColumnAlias[i];
                            oUDO.FindColumns.ColumnDescription = FindColumnDescription[i];
                            if (i < FindColumnAlias.Count() - 1)
                                oUDO.FindColumns.Add();
                        }
                    }

                    if (FormColumnAlias != null)
                    {
                        for (int i = 0; i <= FormColumnAlias.Count() - 1; i++)
                        {
                            oUDO.FormColumns.FormColumnAlias = FormColumnAlias[i];
                            oUDO.FormColumns.FormColumnDescription = FormColumnDescription[i];
                            oUDO.FormColumns.Editable = FormEditable[i];
                            if (i < FormColumnAlias.Count() - 1)
                                oUDO.FormColumns.Add();
                        }
                    }

                    if (ChildTableName != null)
                    {
                        oUDO.ChildTables.SetCurrentLine(0);
                        oUDO.ChildTables.TableName = ChildTableName;

                        if (EnhancedFormColumnAlias != null)
                        {
                            oUDO.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES;
                            for (int i = 0; i <= EnhancedFormColumnAlias.Count() - 1; i++)
                            {
                                oUDO.EnhancedFormColumns.ColumnAlias = EnhancedFormColumnAlias[i];
                                oUDO.EnhancedFormColumns.ColumnDescription = EnhancedFormColumnDescription[i];
                                oUDO.EnhancedFormColumns.ColumnIsUsed = EnhancedFormColumnIsUsed[i];
                                oUDO.EnhancedFormColumns.Editable = EnhancedFormEditable[i];
                                oUDO.EnhancedFormColumns.ChildNumber = 1;
                                if (i < EnhancedFormColumnAlias.Count() - 1)
                                    oUDO.EnhancedFormColumns.Add();
                            }
                        }
                        if (ChildTableName2 != null)
                        {
                            oUDO.ChildTables.Add();
                            oUDO.ChildTables.SetCurrentLine(1);
                            oUDO.ChildTables.TableName = ChildTableName2;

                            if (EnhancedFormColumnAlias2 != null)
                            {
                                oUDO.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES;
                                for (int i = 0; i <= EnhancedFormColumnAlias2.Count() - 1; i++)
                                {
                                    oUDO.EnhancedFormColumns.ColumnAlias = EnhancedFormColumnAlias2[i];
                                    oUDO.EnhancedFormColumns.ColumnDescription = EnhancedFormColumnDescription2[i];
                                    oUDO.EnhancedFormColumns.ColumnIsUsed = EnhancedFormColumnIsUsed2[i];
                                    oUDO.EnhancedFormColumns.Editable = EnhancedFormEditable2[i];
                                    oUDO.EnhancedFormColumns.ChildNumber = 2;
                                    if (i < EnhancedFormColumnAlias2.Count() - 1)
                                        oUDO.EnhancedFormColumns.Add();
                                }
                            }
                        }
                    }

                    ret = oUDO.Add();
                    if (ret != 0)
                    {
                        oCompany.GetLastError(out ErrCode, out ErrMsg);
                        oSBOApplication.MessageBox(ErrCode.ToString() + " : " + ErrMsg);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO);
                    oUDO = null/* TODO Change to default(_) if this is not a reference type */;
                    GC.Collect();
                }
            }
            catch (Exception)
            {
                oCompany.GetLastError(out ErrCode, out ErrMsg);
                oSBOApplication.MessageBox(ErrCode.ToString() + " : " + ErrMsg);
            }
            finally
            {
                oRec = null/* TODO Change to default(_) if this is not a reference type */;
                oUDO = null/* TODO Change to default(_) if this is not a reference type */;
                GC.Collect();
            }
        }

        public static void CreateUDOTemplate(SAPbobsCOM.Company oCompany,
                                        string UDOCode, string UDOName
                                        , SAPbobsCOM.BoUDOObjType ObjectType
                                        , string TableName
                                        , SAPbobsCOM.BoYesNoEnum CanFind
                                        , SAPbobsCOM.BoYesNoEnum CanDelete
                                        , SAPbobsCOM.BoYesNoEnum CanCancel
                                        , SAPbobsCOM.BoYesNoEnum CanClose
                                        , SAPbobsCOM.BoYesNoEnum ManageSeries
                                        , SAPbobsCOM.BoYesNoEnum CanYearTransfer
                                        , SAPbobsCOM.BoYesNoEnum CanArchive
                                        , SAPbobsCOM.BoYesNoEnum CanCreateDefaultForm
                                        , SAPbobsCOM.BoYesNoEnum EnableEnhancedForm
                                        , SAPbobsCOM.BoYesNoEnum CanLog
                                        , string LogTableName
                                        , string MenuUID = null
                                        , string MenuCaption = null
                                        , int FatherMenuID = default(int)
                                        , int Position = default(int)
                                        , string[] FindColumnAlias = null
                                        , string[] FindColumnDescription = null
                                        , string[] FormColumnAlias = null
                                        , string[] FormColumnDescription = null
                                        , SAPbobsCOM.BoYesNoEnum[] FormEditable = null
                                        , UDOChild[] ChildsTables = null)
        {
            bool Exists = false;

            #region Check Existing UDO
            SAPbobsCOM.Recordset oRec = null;
            try
            {
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery("SELECT \"Code\" FROM OUDO WHERE \"Code\" = '" + UDOCode + "'");

                if (oRec.RecordCount > 0)
                {
                    Exists = true;
                }
            }
            finally
            {
                Utils.releaseObject(oRec);
            }
            #endregion

            if (!Exists)
            {
                SAPbobsCOM.UserObjectsMD oUDO = null;
                int lErrCode = 0;
                string sErrMsg = "";

                try
                {
                    oUDO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

                    oUDO.Code = UDOCode;
                    oUDO.Name = UDOName;
                    oUDO.ObjectType = ObjectType;
                    oUDO.TableName = TableName;

                    oUDO.CanFind = CanFind;
                    oUDO.CanDelete = CanDelete;
                    oUDO.CanCancel = CanCancel;
                    oUDO.CanClose = CanClose;
                    oUDO.ManageSeries = ManageSeries;
                    oUDO.CanYearTransfer = CanYearTransfer;
                    oUDO.CanArchive = CanArchive;

                    if (CanLog == SAPbobsCOM.BoYesNoEnum.tYES)
                    {
                        oUDO.CanLog = CanLog;
                        oUDO.LogTableName = LogTableName;
                    }

                    if (CanCreateDefaultForm == SAPbobsCOM.BoYesNoEnum.tYES)
                    {
                        oUDO.CanCreateDefaultForm = CanCreateDefaultForm;
                        oUDO.EnableEnhancedForm = EnableEnhancedForm;
                    }
                    else
                    {
                        oUDO.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
                    }

                    if (FindColumnAlias != null)
                    {
                        for (int i = 0; i <= FindColumnAlias.Length - 1; i++)
                        {
                            oUDO.FindColumns.ColumnAlias = FindColumnAlias[i];
                            oUDO.FindColumns.ColumnDescription = FindColumnDescription[i];
                            if (i < FindColumnAlias.Length - 1)
                                oUDO.FindColumns.Add();
                        }
                    }

                    if (FormColumnAlias != null)
                    {
                        for (int i = 0; i <= FormColumnAlias.Length - 1; i++)
                        {
                            oUDO.FormColumns.FormColumnAlias = FormColumnAlias[i];
                            oUDO.FormColumns.FormColumnDescription = FormColumnDescription[i];
                            oUDO.FormColumns.Editable = FormEditable[i];
                            if (i < FormColumnAlias.Length - 1)
                                oUDO.FormColumns.Add();
                        }
                    }

                    if (MenuUID != null)
                    {
                        oUDO.MenuUID = MenuUID;
                        oUDO.MenuItem = SAPbobsCOM.BoYesNoEnum.tYES;
                        oUDO.MenuCaption = MenuCaption;
                        oUDO.FatherMenuID = FatherMenuID;
                        oUDO.Position = Position;
                    }
                    else
                        oUDO.MenuItem = SAPbobsCOM.BoYesNoEnum.tNO;

                    if (ChildsTables != null)
                    {
                        for (int i = 0; i < ChildsTables.Length; i++)
                        {
                            SAPbobsCOM.UserObjectMD_ChildTables oChild = oUDO.ChildTables;
                            try
                            {
                                oChild.SetCurrentLine(i);
                                oChild.TableName = ChildsTables[i].getChildTable();
                                oChild.Add();
                            }
                            finally
                            {
                                Utils.releaseObject(oChild);
                            }

                            if (ChildsTables[i].getEnhancedFormColumnAlias() != null)
                            {
                                oUDO.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES;
                                for (int j = 0; j <= ChildsTables[i].getEnhancedFormColumnAlias().Length - 1; j++)
                                {
                                    oUDO.EnhancedFormColumns.ColumnAlias = ChildsTables[i].getEnhancedFormColumnAlias()[j];
                                    oUDO.EnhancedFormColumns.ColumnDescription = ChildsTables[i].getEnhancedFormColumnDescription()[j];
                                    oUDO.EnhancedFormColumns.ColumnIsUsed = ChildsTables[i].getEnhancedFormColumnIsUsed()[j];
                                    oUDO.EnhancedFormColumns.Editable = ChildsTables[i].getEnhancedFormEditable()[j];
                                    oUDO.EnhancedFormColumns.ChildNumber = i + 1;
                                    if (j < ChildsTables[i].getEnhancedFormColumnAlias().Length - 1)
                                        oUDO.EnhancedFormColumns.Add();
                                }
                            }
                        }
                    }

                    lErrCode = oUDO.Add();

                    if (lErrCode != 0) throw new Exception();
                }
                catch
                {
                    oCompany.GetLastError(out lErrCode, out sErrMsg);
                }
                finally
                {
                    Utils.releaseObject(oUDO);
                }
            }

        }

        public class UDOChild
        {
            protected string ChildTableName;
            protected string[] EnhancedFormColumnAlias;
            protected string[] EnhancedFormColumnDescription;
            protected SAPbobsCOM.BoYesNoEnum[] EnhancedFormColumnIsUsed;
            protected SAPbobsCOM.BoYesNoEnum[] EnhancedFormEditable;

            public UDOChild(string uChildTableName = null
                         , string[] uEnhancedFormColumnAlias = null
                         , string[] uEnhancedFormColumnDescription = null
                         , SAPbobsCOM.BoYesNoEnum[] uEnhancedFormColumnIsUsed = null
                         , SAPbobsCOM.BoYesNoEnum[] uEnhancedFormEditable = null)
            {
                ChildTableName = uChildTableName;
                EnhancedFormColumnAlias = uEnhancedFormColumnAlias;
                EnhancedFormColumnDescription = uEnhancedFormColumnDescription;
                EnhancedFormColumnIsUsed = uEnhancedFormColumnIsUsed;
                EnhancedFormEditable = uEnhancedFormEditable;
            }

            public string getChildTable()
            {
                return ChildTableName;
            }

            public string[] getEnhancedFormColumnAlias()
            {
                return EnhancedFormColumnAlias;
            }

            public string[] getEnhancedFormColumnDescription()
            {
                return EnhancedFormColumnDescription;
            }

            public SAPbobsCOM.BoYesNoEnum[] getEnhancedFormColumnIsUsed()
            {
                return EnhancedFormColumnIsUsed;
            }

            public SAPbobsCOM.BoYesNoEnum[] getEnhancedFormEditable()
            {
                return EnhancedFormEditable;
            }
        }

        public static void CreateQueryCategory(SAPbobsCOM.Company oSBOCompany, string CategoryName)
        {
            CreateQueryCategory(oSBOCompany, CategoryName, "YYYYYYYYYYYYYYY");
        }

        public static void CreateQueryCategory(SAPbobsCOM.Company oSBOCompany, string CategoryName, string Permissions)
        {
            SAPbobsCOM.QueryCategories oQueryCategory;
            SAPbobsCOM.Recordset oRec;
            int lRetCode = 0;
            int lErrCode = 0;
            string sErrMsg = "";
            bool bIsContinue = true;

            oRec = (SAPbobsCOM.Recordset)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec.DoQuery(GeneralVariables.SQLHandler.CheckQueryCategorySQL(CategoryName));

            if (oRec.RecordCount == 0)
                bIsContinue = true;
            else
                bIsContinue = false;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            oRec = null;
            GC.Collect();

            if (bIsContinue)
            {
                oQueryCategory = (SAPbobsCOM.QueryCategories)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories);
                oQueryCategory.Name = CategoryName;
                oQueryCategory.Permissions = Permissions;

                lRetCode = oQueryCategory.Add();
                if (lRetCode != 0)
                {
                    lErrCode = oSBOCompany.GetLastErrorCode();
                    sErrMsg = oSBOCompany.GetLastErrorDescription();
                    throw new Exception(lErrCode.ToString() + " : " + sErrMsg);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oQueryCategory);
                oQueryCategory = null;
                GC.Collect();
            }
        }

        public static void CreateQuery(SAPbobsCOM.Company oSBOCompany, string CategoryName, string QueryName, string SQL)
        {
            SAPbobsCOM.UserQueries oUserQuery;
            SAPbobsCOM.Recordset oRec;
            int lRetCode = 0;
            int lErrCode = 0;
            string sErrMsg = "";
            bool bIsContinue = true;
            int CategoryId = 0;

            oRec = (SAPbobsCOM.Recordset)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec.DoQuery(GeneralVariables.SQLHandler.CheckQueryCategorySQL(CategoryName));

            if (oRec.RecordCount != 0)
                CategoryId = System.Convert.ToInt32(oRec.Fields.Item("CategoryId").Value);
            else
                throw new Exception("Query Category (" + CategoryName + ") Not Exists!!!");

            oRec.DoQuery(GeneralVariables.SQLHandler.CheckQuerySQL(CategoryName, QueryName));

            if (oRec.RecordCount == 0)
                bIsContinue = true;
            else
                bIsContinue = false;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            oRec = null;
            GC.Collect();

            if (bIsContinue)
            {
                oUserQuery = (SAPbobsCOM.UserQueries)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries);
                oUserQuery.Query = SQL;
                oUserQuery.QueryCategory = CategoryId;
                oUserQuery.QueryDescription = QueryName;

                lRetCode = oUserQuery.Add();
                if (lRetCode != 0)
                {
                    lErrCode = oSBOCompany.GetLastErrorCode();
                    sErrMsg = oSBOCompany.GetLastErrorDescription();
                    throw new Exception(lErrCode.ToString() + " : " + sErrMsg);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserQuery);
                oUserQuery = null;
                GC.Collect();
            }
        }

        public static void CreateQueryFromText(SAPbobsCOM.Company oSBOCompany, Application oSBOApplication, string queryCat, string queryName, string textFile)
        {
            try
            {
                string dir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string dirPathFMS = dir + @"\FMS\";

                string path = Path.Combine(dirPathFMS, Path.GetFileName(textFile));
                string readText = File.ReadAllText(path);

                if (!string.IsNullOrEmpty(readText))
                {
                    CreateQuery(oSBOCompany, queryCat, queryName, readText);
                }
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
        }

        public static void CreateFMS(SAPbobsCOM.Company oSBOCompany
                                    , string CategoryName
                                    , string QueryName
                                    , string FormID
                                    , string ItemID
                                    , BoYesNoEnum autoRefresh
                                    , BoYesNoEnum refreshByField
                                    , BoYesNoEnum forceRefresh
                                    , string ColID = null
                                    , string userField = null)
        {
            SAPbobsCOM.FormattedSearches oFormattedSearch;
            SAPbobsCOM.Recordset oRec;
            int lRetCode = 0;
            int lErrCode = 0;
            string sErrMsg = "";
            bool bIsContinue = true;
            int QueryId;
            int CategoryId;

            oRec = (SAPbobsCOM.Recordset)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oRec.DoQuery(GeneralVariables.SQLHandler.CheckQueryCategorySQL(CategoryName));
            if (oRec.RecordCount != 0)
                CategoryId = System.Convert.ToInt32(oRec.Fields.Item("CategoryId").Value);
            else
                throw new Exception("Query Category (" + CategoryName + ") Not Exists!!!");

            oRec.DoQuery(GeneralVariables.SQLHandler.CheckQuerySQL(CategoryName, QueryName));
            if (oRec.RecordCount != 0)
                QueryId = System.Convert.ToInt32(oRec.Fields.Item("IntrnalKey").Value);
            else
                throw new Exception("Query (" + QueryName + ") Not Exists!!!");

            oRec.DoQuery(GeneralVariables.SQLHandler.CheckFMSSQL(FormID, ItemID, ColID));

            if (oRec.RecordCount == 0)
                bIsContinue = true;
            else
                bIsContinue = false;

            //string IndexID = oRec.Fields.Item("IndexID").Value.ToString.Trim;
            int IndexID = oRec.Fields.Item("IndexID").Value;

            //string resultQueryID = oRec.Fields.Item("QueryId").Value.ToString.Trim;
            int resultQueryID = oRec.Fields.Item("QueryId").Value;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            oRec = null;
            GC.Collect();

            if (bIsContinue)
            {
                oFormattedSearch = (SAPbobsCOM.FormattedSearches)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
                oFormattedSearch.FormID = FormID;
                oFormattedSearch.ItemID = ItemID;
                oFormattedSearch.ColumnID = ColID;
                oFormattedSearch.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
                oFormattedSearch.QueryID = QueryId;
                oFormattedSearch.Refresh = autoRefresh;
                oFormattedSearch.ForceRefresh = forceRefresh;
                oFormattedSearch.ByField = refreshByField;
                oFormattedSearch.FieldID = userField;
                lRetCode = oFormattedSearch.Add();

                if (lRetCode != 0)
                {
                    lErrCode = oSBOCompany.GetLastErrorCode();
                    sErrMsg = oSBOCompany.GetLastErrorDescription();
                    throw new Exception(lErrCode.ToString() + " : " + sErrMsg);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oFormattedSearch);
                oFormattedSearch = null;
                GC.Collect();
            }
            else if (QueryId != resultQueryID)
            {
                oFormattedSearch = (SAPbobsCOM.FormattedSearches)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
                if (oFormattedSearch.GetByKey(IndexID))
                {
                    oFormattedSearch.FormID = FormID;
                    oFormattedSearch.ItemID = ItemID;
                    oFormattedSearch.ColumnID = ColID;
                    oFormattedSearch.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
                    oFormattedSearch.QueryID = QueryId;
                    oFormattedSearch.ForceRefresh = forceRefresh;
                    oFormattedSearch.Refresh = autoRefresh;
                    oFormattedSearch.ByField = refreshByField;
                    oFormattedSearch.FieldID = userField;

                    lRetCode = oFormattedSearch.Update();

                    if (lRetCode != 0)
                    {
                        lErrCode = oSBOCompany.GetLastErrorCode();
                        sErrMsg = oSBOCompany.GetLastErrorDescription();
                        throw new Exception(lErrCode.ToString() + " : " + sErrMsg);
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oFormattedSearch);
                    oFormattedSearch = null;
                    GC.Collect();
                }
            }
        }

        public static void CreateSP(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplication
                                    , string fileName, string spName)
        {
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                Queries queries = null;
                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                {
                    queries = new HANAQueries();
                    oRec.DoQuery(queries.CheckSPExistsSQL(oSBOCompany.CompanyDB, spName));
                }
                else
                {
                    queries = new SQLQueries();
                    oRec.DoQuery(queries.CheckSPExistsSQL(oSBOCompany.CompanyDB, spName));
                }

                if (oRec.RecordCount <= 0)
                {
                    string dir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                    string dirPathSP = dir + @"\SP\";
                    string directory = Path.Combine(dirPathSP, Path.GetFileName(fileName));

                    using (StreamReader sr = new StreamReader(directory))
                    {
                        string line = sr.ReadToEnd();
                        oRec.DoQuery(line);
                    }
                }

            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
            finally
            {
                releaseObject(oRec);
            }
        }

        // 20140225
        public static DateTime ConvertToDateTime(string SBODate)
        {
            return new DateTime(int.Parse(SBODate.Substring(0, 4)), int.Parse(SBODate.Substring(4, 2)), int.Parse(SBODate.Substring(6, 2)));
        }

        public static string WindowsToSBONumber(double value)
        {
            return value.ToString().Replace(GeneralVariables.WinDecSep, GeneralVariables.SBODecSep);
        }

        public static double SBOToWindowsNumberWithCurrency(string value)
        {
            return SBOToWindowsNumberWithoutCurrency(value.Substring(4).ToString());
        }

        public static double SBOToWindowsNumberWithoutCurrency(string value)
        {
            return double.Parse(value.Replace(GeneralVariables.SBOThousSep, "").Replace(GeneralVariables.SBODecSep, GeneralVariables.WinDecSep));
        }

        public static string FormattedStringAmount(double value)
        {
            string valueS = value.ToString("G17");
            int DecSepIndex = valueS.Trim().IndexOf(GeneralVariables.WinDecSep);

            string a = valueS;
            string b = "";

            if (DecSepIndex >= 0)
            {
                a = valueS.Trim().Substring(0, valueS.Trim().IndexOf(GeneralVariables.WinDecSep));
                b = valueS.Trim().Substring(valueS.Trim().IndexOf(GeneralVariables.WinDecSep) + 1);
            }

            int c = (int)Math.Floor(double.Parse(a.Length.ToString()) / double.Parse("3"));

            string d = StringReverse(a);

            List<string> e = new List<string>();

            int ctr = 0;
            for (int i = 0; i <= c - 1; i++)
            {
                string f = d.Substring(i + ctr, 3);

                ctr = ctr + 2;

                e.Add(StringReverse(f));
            }

            string g = "";
            string h = "";

            for (int i = e.Count - 1; i >= 0; i += -1)
            {
                g = g + e.ElementAt(i).ToString();
                h = h + e.ElementAt(i).ToString() + GeneralVariables.SBOThousSep;
            }

            string result = a.Substring(0, a.Length - g.Length) + GeneralVariables.SBOThousSep + h;
            result = result.Substring(0, result.Length - 1);

            if (result.StartsWith(GeneralVariables.SBOThousSep))
                result = result.Substring(1);

            if (b != "")
                result = result + GeneralVariables.SBODecSep + b;

            return result;
        }

        public static System.Data.DataTable ConvertRecordsetToDataTable(SAPbobsCOM.Company oSBOCompany, string sql)
        {
            SAPbobsCOM.Recordset oRec;
            oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oRec.DoQuery(sql);

            return ConvertRecordsetToDataTable(oRec);
        }

        public static System.Data.DataTable ConvertRecordsetToDataTable(SAPbobsCOM.Recordset oRec)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            for (int i = 0; i <= oRec.Fields.Count - 1; i++)
                dt.Columns.Add(oRec.Fields.Item(i).Name);

            while (!oRec.EoF)
            {
                List<object> innerDataList = new List<object>();

                for (int i = 0; i <= oRec.Fields.Count - 1; i++)
                    innerDataList.Add(oRec.Fields.Item(i).Value);

                dt.Rows.Add(innerDataList.ToArray());
                oRec.MoveNext();
            }

            return dt;
        }

        public static string StringReverse(string s)
        {
            char[] charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }

        public static Form createForm(ref Application oApp, string srfName)
        {
            XmlDocument oXMLDoc = new XmlDocument();

            SAPbouiCOM.FormCreationParams oCreationPackage = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
            oCreationPackage.UniqueID = String.Format("{0}{1}", oCreationPackage.UniqueID, Guid.NewGuid().ToString().Substring(2, 10)).Replace("-", String.Empty);

            string dir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = dir + @"\SRF\";

            //string path = "../../SRF/";
            oXMLDoc.Load(path + srfName + ".srf");
            oCreationPackage.XmlData = oXMLDoc.InnerXml;
            oCreationPackage.BorderStyle = BoFormBorderStyle.fbs_Sizable;

            return oApp.Forms.AddEx(oCreationPackage);
        }

        public static void CreateMenuEntryTemplate(ref Application oApp, ref SAPbobsCOM.Company oCompany, string menuUID
                                                , string caption
                                                , SAPbouiCOM.BoMenuType menuType
                                                , int position = 0
                                                , string parentMenuUID = null
                                                , string imgPath = null)
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;
            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            int lErrCode = 0;
            string sErrMsg = "";

            try
            {
                oMenus = oApp.Menus;

                if (!oMenus.Exists(menuUID))
                {
                    oCreationPackage = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                    if (parentMenuUID != null || oMenus.Exists(parentMenuUID)) oMenuItem = oApp.Menus.Item(parentMenuUID);
                    else oMenuItem = oApp.Menus.Item("43520");

                    oMenus = oMenuItem.SubMenus;

                    oCreationPackage.UniqueID = menuUID;
                    oCreationPackage.String = caption;
                    oCreationPackage.Type = menuType;

                    if (position > 0) oCreationPackage.Position = position;
                    else oCreationPackage.Position = oMenus.Count + 1;

                    if (imgPath != null) oCreationPackage.Image = imgPath;
                    else oCreationPackage.Image = "";

                    oMenus.AddEx(oCreationPackage);
                }
            }
            catch
            {
                oCompany.GetLastError(out lErrCode, out sErrMsg);
                oApp.MessageBox(lErrCode.ToString() + " : " + sErrMsg);
            }
            finally
            {
                releaseObject(oMenus);
                releaseObject(oMenuItem);
                releaseObject(oCreationPackage);
            }
        }

        public static void releaseObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch { obj = null; }
            finally { GC.Collect(); }
        }
    }
}
