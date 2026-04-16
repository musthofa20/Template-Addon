using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.IO;

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
        CreateUDF(oSBOCompany, TableName, FieldName, FieldDescription, Type, SubType, EditSize, ValidValues, DefaultValue, SAPbobsCOM.BoYesNoEnum.tNO);
    }

    public static void CreateUDF(SAPbobsCOM.Company oSBOCompany, string TableName, string FieldName, string FieldDescription, SAPbobsCOM.BoFieldTypes Type, SAPbobsCOM.BoFldSubTypes SubType, int EditSize, List<string[]> ValidValues, string DefaultValue, SAPbobsCOM.BoYesNoEnum Mandatory)
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

    public static void CreateUDF(SAPbobsCOM.Company oSBOCompany, string TableName, string FieldName, string FieldDescription, SAPbobsCOM.BoFieldTypes Type, SAPbobsCOM.BoFldSubTypes SubType, int EditSize, string LinkedTable, string DefaultValue, SAPbobsCOM.BoYesNoEnum Mandatory)
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

    public static void CreateMenu(SAPbobsCOM.Company oSBOCompany, SAPbouiCOM.Application oSBOApplication, string ParentMenuUID, string MenuUID, SAPbouiCOM.BoMenuType MenuType, string MenuName, string MenuImage, int MenuPosition, bool DeleteIfExists)
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

    public static void CreateUDO(SAPbobsCOM.Company oSBOCompany, string ObjectName, SAPbobsCOM.BoUDOObjType ObjectType, string TableName, SAPbobsCOM.BoYesNoEnum CanApprove, SAPbobsCOM.BoYesNoEnum CanArchive, SAPbobsCOM.BoYesNoEnum CanCancel, SAPbobsCOM.BoYesNoEnum CanClose, SAPbobsCOM.BoYesNoEnum CanCreateDefaultForm, SAPbobsCOM.BoYesNoEnum CanDelete, SAPbobsCOM.BoYesNoEnum CanFind, SAPbobsCOM.BoYesNoEnum CanLog, SAPbobsCOM.BoYesNoEnum CanYearTransfer, SAPbobsCOM.BoYesNoEnum ManageSeries, string[] FindColumns, string[] ChildTables)
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

    public static void CreateUDO(SAPbobsCOM.Company oSBOCompany, string ObjectName, SAPbobsCOM.BoUDOObjType ObjectType, string TableName, SAPbobsCOM.BoYesNoEnum CanApprove, SAPbobsCOM.BoYesNoEnum CanArchive, SAPbobsCOM.BoYesNoEnum CanCancel, SAPbobsCOM.BoYesNoEnum CanClose, SAPbobsCOM.BoYesNoEnum CanCreateDefaultForm, SAPbobsCOM.BoYesNoEnum CanDelete, SAPbobsCOM.BoYesNoEnum CanFind, SAPbobsCOM.BoYesNoEnum CanLog, SAPbobsCOM.BoYesNoEnum CanYearTransfer, SAPbobsCOM.BoYesNoEnum ManageSeries, string FatherMenuID, string MenuCaption, string[] FindColumns, string[] ChildTables, string[,] FormColumns)
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

            oUserObjectMD.EnableEnhancedForm = CanCreateDefaultForm;
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

    public static void CreateFMS(SAPbobsCOM.Company oSBOCompany, string CategoryName, string QueryName, string FormID, string ItemID)
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

        oRec.DoQuery(GeneralVariables.SQLHandler.CheckFMSSQL(FormID, ItemID));

        if (oRec.RecordCount == 0)
            bIsContinue = true;
        else
            bIsContinue = false;

        string IndexID;
        IndexID = oRec.Fields.Item("IndexID").Value.ToString.Trim;

        string resultQueryID;
        resultQueryID = oRec.Fields.Item("QueryId").Value.ToString.Trim;

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
        oRec = null;
        GC.Collect();

        if (bIsContinue)
        {
            oFormattedSearch = (SAPbobsCOM.FormattedSearches)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
            oFormattedSearch.FormID = FormID;
            oFormattedSearch.ItemID = ItemID;
            oFormattedSearch.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
            oFormattedSearch.QueryID = QueryId;

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
        else if (QueryId.ToString() != resultQueryID)
        {
            oFormattedSearch = (SAPbobsCOM.FormattedSearches)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
            if (oFormattedSearch.GetByKey(int.Parse(IndexID)))
            {
                oFormattedSearch.FormID = FormID;
                oFormattedSearch.ItemID = ItemID;
                oFormattedSearch.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
                oFormattedSearch.QueryID = QueryId;

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

    public static void ExecFunction(SAPbobsCOM.Company oSBOCompany, string FileName, string FunctionName)
    {
        SAPbobsCOM.Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        oRec.DoQuery(GeneralVariables.SQLHandler.CheckFunctionExistsSQL(oSBOCompany.CompanyDB, FunctionName));

        if (oRec.RecordCount <= 0)
        {
            StreamReader sr = new StreamReader(FileName);
            string line = sr.ReadToEnd();
            oRec.DoQuery(line);
        }

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
        oRec = null;
        GC.Collect();
    }

    public static void ExecSP(SAPbobsCOM.Company oSBOCompany, string FileName, string SPName)
    {
        SAPbobsCOM.Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        oRec.DoQuery(GeneralVariables.SQLHandler.CheckSPExistsSQL(oSBOCompany.CompanyDB, SPName));

        if (oRec.RecordCount <= 0)
        {
            StreamReader sr = new StreamReader(FileName);
            string line = sr.ReadToEnd();
            oRec.DoQuery(line);
        }

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
        oRec = null;
        GC.Collect();
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

    public static DataTable ConvertRecordsetToDataTable(SAPbobsCOM.Company oSBOCompany, string sql)
    {
        SAPbobsCOM.Recordset oRec;
        oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        oRec.DoQuery(sql);

        return ConvertRecordsetToDataTable(oRec);
    }

    public static DataTable ConvertRecordsetToDataTable(SAPbobsCOM.Recordset oRec)
    {
        DataTable dt = new DataTable();

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


    //DECIMAL & DOUBLE TO SBO NUMBER
    public static string DecimalToSBONumber(Decimal value)
    {
        string RetVal;
        string WindowsDecimalSeparator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        RetVal = Convert.ToString(value).Replace(Convert.ToChar(WindowsDecimalSeparator), '.');
        return RetVal;
    }

    public static string DoubleToSBONumber(double value)
    {
        string RetVal;
        string WindowsDecimalSeparator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        RetVal = Convert.ToString(value).Replace(',', '.');
        return RetVal;
    }

    //SBO NUMBER TO DECIMAL & DOUBLE
    public static Decimal SBONumbertoDecimal(string value)
    {
        decimal RetVal;
        string WindowsDecimalSeparator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        RetVal = Convert.ToDecimal(value.Replace(".", WindowsDecimalSeparator));
        return RetVal;
    }

    public static Double SBONumbertoDouble(string value)
    {
        Double RetVal;
        RetVal = double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
        return RetVal;
    }

}
