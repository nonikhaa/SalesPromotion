public class SQLQueries : Queries
{

    public string SeparatorSQL()
    {
        return "SELECT DecSep, ThousSep FROM OADM";
    }

    public string CheckUDFSQL(string tableName, string fieldName)
    {
        return "SELECT TableID FROM CUFD WHERE TableID = '" + tableName.Trim() + "' AND AliasID = '" + fieldName.Trim() + "'";
    }

    public string CheckUDOSQL(string tableName)
    {
        return "SELECT Code FROM OUDO WHERE Code = '" + tableName.Trim() + "'";
    }

    public string CheckQueryCategorySQL(string categoryName)
    {
        return "SELECT CategoryId FROM OQCN WHERE CatName = '" + categoryName + "'";
    }

    public string CheckQuerySQL(string categoryName, string queryName)
    {
        return "SELECT IntrnalKey FROM OUQR WHERE QName = '" + queryName + "' AND QCategory IN (" + CheckQueryCategorySQL(categoryName) + ")";
    }

    public string CheckFMSSQL(string formID, string itemID, string columnID)
    {
        return "SELECT IndexID, QueryId FROM CSHS WHERE FormID = '" + formID + "' AND ItemID = '" + itemID + "' AND ISNULL(ColID, '') = '" + columnID + "'";
    }

    public string CheckFunctionExistsSQL(string DBName, string FunctionName)
    {
        return "SELECT TOP 1 1 FROM INFORMATION_SCHEMA.ROUTINES WHERE SPECIFIC_CATALOG = '" + DBName + "' AND SPECIFIC_SCHEMA = 'dbo' AND SPECIFIC_NAME = '" + FunctionName + "' AND ROUTINE_TYPE = 'FUNCTION'";
    }

    public string CheckSPExistsSQL(string DBName, string SPName)
    {
        return "SELECT TOP 1 1 FROM INFORMATION_SCHEMA.ROUTINES WHERE SPECIFIC_CATALOG = '" + DBName + "' AND SPECIFIC_SCHEMA = 'dbo' AND SPECIFIC_NAME = '" + SPName + "' AND ROUTINE_TYPE = 'PROCEDURE'";
    }

}
