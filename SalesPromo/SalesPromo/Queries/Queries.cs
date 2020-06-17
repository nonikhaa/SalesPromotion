public interface Queries
{

    string SeparatorSQL();
    string CheckUDFSQL(string tableName, string fieldName);
    string CheckUDOSQL(string tableName);
    string CheckQueryCategorySQL(string categoryName);
    string CheckQuerySQL(string categoryName, string queryName);
    string CheckFMSSQL(string formID, string itemID, string columnID);
    string CheckFunctionExistsSQL(string DBName, string FunctionName);
    string CheckSPExistsSQL(string DBName, string SPName);

}
