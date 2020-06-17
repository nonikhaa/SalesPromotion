public class HANAQueries : Queries
{

    public string SeparatorSQL()
    {
        return "SELECT \"DecSep\", \"ThousSep\" FROM \"OADM\"";
    }

    public string CheckUDFSQL(string tableName, string fieldName)
    {
        return "SELECT \"TableID\" FROM \"CUFD\" WHERE \"TableID\" = '" + tableName.Trim() + "' AND \"AliasID\" = '" + fieldName.Trim() + "'";
    }

    public string CheckUDOSQL(string tableName)
    {
        return "SELECT \"Code\" FROM \"OUDO\" WHERE \"Code\" = '" + tableName.Trim() + "'";
    }

    public string CheckQueryCategorySQL(string categoryName)
    {
        return "SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\" = '" + categoryName + "'";
    }

    public string CheckQuerySQL(string categoryName, string queryName)
    {
        return "SELECT \"IntrnalKey\" FROM \"OUQR\" WHERE \"QName\" = '" + queryName + "' AND \"QCategory\" IN (" + CheckQueryCategorySQL(categoryName) + ")";
    }

    public string CheckFMSSQL(string formID, string itemID, string columnID)
    {
        return "SELECT \"IndexID\", \"QueryId\" FROM \"CSHS\" WHERE \"FormID\" = '" + formID + "' AND \"ItemID\" = '" + itemID + "' AND IFNULL(\"ColID\", '') = '" + columnID + "'";
    }

    public string CheckFunctionExistsSQL(string DBName, string FunctionName)
    {
        return "SELECT TOP 1 1 FROM OBJECTS WHERE SCHEMA_NAME = '" + DBName + "' AND OBJECT_TYPE = 'FUNCTION' AND OBJECT_NAME = '" + FunctionName + "'";
    }

    public string CheckSPExistsSQL(string DBName, string SPName)
    {
        return "SELECT TOP 1 1 FROM OBJECTS WHERE SCHEMA_NAME = '" + DBName + "' AND OBJECT_TYPE = 'PROCEDURE' AND OBJECT_NAME = '" + SPName + "'";
    }

}
