namespace MsAccessExtract.Helpers;

/// <summary>
/// Access and VBA constants for COM late binding.
/// These mirror the enum values from Microsoft.Office.Interop.Access and VBIDE.
/// </summary>
internal static class AccessConstants
{
    // AcObjectType — used by SaveAsText / LoadFromText
    public const int AcTable = 0;
    public const int AcQuery = 1;
    public const int AcForm = 2;
    public const int AcReport = 3;
    public const int AcMacro = 4;
    public const int AcModule = 5;
    public const int AcDatabaseProperties = 11;

    // vbext_ComponentType — VBComponent.Type
    public const int VbextCtStdModule = 1;
    public const int VbextCtClassModule = 2;
    public const int VbextCtMsForm = 3;
    public const int VbextCtDocument = 100;

    // DAO DataTypeEnum — TableDef.Fields.Type
    public const int DbBoolean = 1;
    public const int DbByte = 2;
    public const int DbInteger = 3;
    public const int DbLong = 4;
    public const int DbCurrency = 5;
    public const int DbSingle = 6;
    public const int DbDouble = 7;
    public const int DbDate = 8;
    public const int DbBinary = 9;
    public const int DbText = 10;
    public const int DbLongBinary = 11;
    public const int DbMemo = 12;
    public const int DbGuid = 15;
    public const int DbBigInt = 16;
    public const int DbVarBinary = 17;
    public const int DbChar = 18;
    public const int DbNumeric = 19;
    public const int DbFloat = 21;
    public const int DbTime = 22;
    public const int DbTimeStamp = 23;
    public const int DbAttachment = 101;
    public const int DbComplexByte = 102;
    public const int DbComplexInteger = 103;
    public const int DbComplexLong = 104;
    public const int DbComplexSingle = 105;
    public const int DbComplexDouble = 106;
    public const int DbComplexGuid = 107;
    public const int DbComplexDecimal = 108;
    public const int DbComplexText = 109;

    // DAO Field Attributes
    public const int DbAutoIncrField = 16;
    public const int DbFixedField = 1;
    public const int DbVariableField = 2;
    public const int DbHyperlinkField = 32768;

    // DAO Relation Attributes
    public const int DbRelationUnique = 1;
    public const int DbRelationDontEnforce = 2;
    public const int DbRelationInherited = 4;
    public const int DbRelationUpdateCascade = 256;
    public const int DbRelationDeleteCascade = 4096;

    public static string GetFieldTypeName(int typeValue) => typeValue switch
    {
        DbBoolean => "Boolean",
        DbByte => "Byte",
        DbInteger => "Integer",
        DbLong => "Long",
        DbCurrency => "Currency",
        DbSingle => "Single",
        DbDouble => "Double",
        DbDate => "DateTime",
        DbBinary => "Binary",
        DbText => "Text",
        DbLongBinary => "OLE Object",
        DbMemo => "Memo",
        DbGuid => "GUID",
        DbBigInt => "BigInt",
        DbVarBinary => "VarBinary",
        DbChar => "Char",
        DbNumeric => "Numeric",
        DbFloat => "Float",
        DbTime => "Time",
        DbTimeStamp => "TimeStamp",
        DbAttachment => "Attachment",
        DbComplexByte => "ComplexByte",
        DbComplexInteger => "ComplexInteger",
        DbComplexLong => "ComplexLong",
        DbComplexSingle => "ComplexSingle",
        DbComplexDouble => "ComplexDouble",
        DbComplexGuid => "ComplexGUID",
        DbComplexDecimal => "ComplexDecimal",
        DbComplexText => "ComplexText",
        _ => $"Unknown({typeValue})"
    };
}
