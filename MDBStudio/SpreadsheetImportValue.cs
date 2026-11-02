using System.Data;
using System.Globalization;

namespace MDBStudio;

/// <summary>
/// Converts imported Excel/CSV field text to the destination Access column type.
/// Access Yes/No columns conventionally encode True as -1 and False as 0.
/// </summary>
internal static class SpreadsheetImportValue
{
    public static object ConvertToColumn(string input, DataColumn column)
    {
        Type type = column.DataType;
        if (type == typeof(string))
        {
            if (column.MaxLength > -1 && input.Length > column.MaxLength)
                throw new FormatException($"Text exceeds the {column.MaxLength}-character limit.");
            return input;
        }
        if (input.Length == 0)
        {
            if (column.AllowDBNull) return DBNull.Value;
            throw new FormatException("The destination column requires a value.");
        }
        if (type == typeof(bool)) return ParseAccessBoolean(input);
        if (type == typeof(DateTime))
            return DateTime.Parse(input, CultureInfo.CurrentCulture);
        if (type == typeof(Guid)) return Guid.Parse(input);
        if (type == typeof(byte[]))
            throw new FormatException("Binary Access columns cannot be imported from a spreadsheet.");
        if (type == typeof(int)) return int.Parse(input, NumberStyles.Integer | NumberStyles.AllowThousands, CultureInfo.CurrentCulture);
        if (type == typeof(long)) return long.Parse(input, NumberStyles.Integer | NumberStyles.AllowThousands, CultureInfo.CurrentCulture);
        if (type == typeof(short)) return short.Parse(input, NumberStyles.Integer | NumberStyles.AllowThousands, CultureInfo.CurrentCulture);
        if (type == typeof(byte)) return byte.Parse(input, NumberStyles.Integer, CultureInfo.CurrentCulture);
        if (type == typeof(double)) return double.Parse(input, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture);
        if (type == typeof(decimal)) return decimal.Parse(input, NumberStyles.Number, CultureInfo.CurrentCulture);
        if (type == typeof(float)) return float.Parse(input, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture);
        throw new FormatException($"The database field type '{type.Name}' is not supported for spreadsheet import.");
    }

    internal static bool ParseAccessBoolean(string text)
    {
        // Jet/ACE stores Yes/No True as -1; exporting a field to CSV may
        // expose -1 even though .NET bool.Parse only accepts True or False.
        switch (text.Trim().ToUpperInvariant())
        {
            case "-1":
            case "1":
            case "TRUE":
            case "YES":
            case "Y":
            case "ON":
            case "T":
                return true;
            case "0":
            case "FALSE":
            case "NO":
            case "N":
            case "OFF":
            case "F":
                return false;
            default:
                throw new FormatException(
                    "Expected a Yes/No value: -1, 0, 1, True, False, Yes or No.");
        }
    }
}
