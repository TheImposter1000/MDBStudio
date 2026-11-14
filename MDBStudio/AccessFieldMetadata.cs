using System.Data;
using System.Data.OleDb;
using System.Globalization;

namespace MDBStudio;

/// <summary>
/// Access field metadata must come from the OLE DB Columns schema, not just
/// DataColumn.MaxLength. FillSchema can report -1 for normal short text fields.
/// </summary>
public static class AccessFieldMetadata
{
    public static string FriendlyType(DataColumn column, int? oleDbType, int? declaredLength)
    {
        if (column.AutoIncrement) return "AutoNumber";

        if (oleDbType is int code)
        {
            if (code is (int)OleDbType.WChar or (int)OleDbType.VarWChar or
                (int)OleDbType.Char or (int)OleDbType.VarChar)
                return "Short Text";
            if (code is (int)OleDbType.LongVarWChar or (int)OleDbType.LongVarChar)
                return "Long Text";
            if (code == (int)OleDbType.Integer) return "Integer (32-bit)";
            if (code == (int)OleDbType.SmallInt) return "Small Integer";
            if (code == (int)OleDbType.UnsignedTinyInt) return "Byte";
            if (code == (int)OleDbType.Single) return "Single";
            if (code == (int)OleDbType.Double) return "Double";
            if (code == (int)OleDbType.Currency) return "Decimal/Currency";
            if (code is (int)OleDbType.Decimal or (int)OleDbType.Numeric)
                return "Decimal/Currency";
            if (code == (int)OleDbType.Boolean) return "Yes/No";
            if (code is (int)OleDbType.Date or (int)OleDbType.DBTimeStamp)
                return "Date/Time";
            if (code == (int)OleDbType.Guid) return "GUID";
            if (code is (int)OleDbType.Binary or (int)OleDbType.VarBinary or
                (int)OleDbType.LongVarBinary)
                return "Binary/OLE";
        }

        Type type = column.DataType;
        if (type == typeof(string))
        {
            int length = declaredLength ?? column.MaxLength;
            return length is > 0 and <= 255 ? "Short Text" : "Long Text";
        }
        if (type == typeof(int)) return "Integer (32-bit)";
        if (type == typeof(short)) return "Small Integer";
        if (type == typeof(byte)) return "Byte";
        if (type == typeof(bool)) return "Yes/No";
        if (type == typeof(DateTime)) return "Date/Time";
        if (type == typeof(double)) return "Double";
        if (type == typeof(float)) return "Single";
        if (type == typeof(decimal)) return "Decimal/Currency";
        if (type == typeof(Guid)) return "GUID";
        if (type == typeof(byte[])) return "Binary/OLE";
        return type.Name;
    }

    public static int TextLength(DataColumn column, int? declaredLength)
    {
        int length = declaredLength ?? column.MaxLength;
        return length > 0 ? length : -1;
    }

    public static bool IsSameType(SchemaColumnInfo existing, AccessFieldDefinition target)
    {
        string name = target.Kind switch
        {
            AccessFieldKind.ShortText => "Short Text",
            AccessFieldKind.LongText => "Long Text",
            AccessFieldKind.Integer => "Integer (32-bit)",
            AccessFieldKind.SmallInteger => "Small Integer",
            AccessFieldKind.Byte => "Byte",
            AccessFieldKind.Double => "Double",
            AccessFieldKind.Currency => "Decimal/Currency",
            AccessFieldKind.DateTime => "Date/Time",
            AccessFieldKind.YesNo => "Yes/No",
            AccessFieldKind.Guid => "GUID",
            AccessFieldKind.AutoNumber => "AutoNumber",
            _ => ""
        };
        return existing.DataType == name &&
            (target.Kind != AccessFieldKind.ShortText || existing.Length == target.Length);
    }

    public static (int? Type, int? Length) ReadColumnMetadata(
        DataTable? columnsSchema, string columnName)
    {
        if (columnsSchema is null || !columnsSchema.Columns.Contains("COLUMN_NAME"))
            return (null, null);

        var row = columnsSchema.Rows.Cast<DataRow>().FirstOrDefault(candidate =>
            string.Equals(Convert.ToString(candidate["COLUMN_NAME"], CultureInfo.InvariantCulture),
                columnName, StringComparison.OrdinalIgnoreCase));
        if (row is null) return (null, null);

        int? ReadInt(string name)
        {
            if (!columnsSchema.Columns.Contains(name) || row[name] is DBNull) return null;
            return Convert.ToInt32(row[name], CultureInfo.InvariantCulture);
        }

        return (ReadInt("DATA_TYPE"), ReadInt("CHARACTER_MAXIMUM_LENGTH"));
    }
}
