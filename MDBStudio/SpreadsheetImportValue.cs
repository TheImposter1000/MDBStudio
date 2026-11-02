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
        // TEW's Victor field is a numeric winner identifier. For angles,
        // -1 is a sentinel meaning "no match winner", not Access True.
        // If an unexpected destination defines Victor as Yes/No we must fail
        // rather than silently corrupt those angle records.
        if (type == typeof(bool) &&
            string.Equals(column.ColumnName, "Victor", StringComparison.OrdinalIgnoreCase) &&
            input.Trim() == "-1")
        {
            throw new FormatException(
                "TEW angle records use Victor = -1 to mean no match winner. " +
                "The destination Victor field is defined as Yes/No (Boolean), " +
                "which cannot preserve the numeric -1 sentinel. " +
                "Import into the correct numeric Victor field/table instead. " +
                "No records have been imported.");
        }

        if (type == typeof(bool)) return ParseAccessBoolean(input);
        if (type == typeof(DateTime))
            return DateTime.Parse(input, CultureInfo.CurrentCulture);
        if (type == typeof(Guid)) return Guid.Parse(input);
        if (type == typeof(byte[]))
            throw new FormatException("Binary Access columns cannot be imported from a spreadsheet.");
        if (type == typeof(int) || type == typeof(long) || type == typeof(short) ||
            type == typeof(byte) || type == typeof(double) || type == typeof(decimal) ||
            type == typeof(float))
            return ParseNumber(input, type);
        throw new FormatException($"The database field type '{type.Name}' is not supported for spreadsheet import.");
    }

    /// <summary>
    /// Normalize only numeric input. Old CSV exports could prefix negative
    /// text numbers with an apostrophe to prevent spreadsheet formulas.
    /// Excel may use Unicode minus signs or insert invisible formatting marks.
    /// These conversions never apply to text/ID columns.
    /// </summary>
    internal static string NormalizeNumericInput(string input)
    {
        string value = input.Trim();
        value = TrimInvisibleBoundary(value);
        if (value.Length > 0)
        {
            char sign = value[0];
            if (sign is '\u2212' or '\uFF0D' or '\u2010' or '\u2011'
                or '\u2012' or '\u2013' or '\uFE63')
                value = "-" + value[1..];
            else if (sign == '\uFF0B')
                value = "+" + value[1..];
        }

        // Text CSV exports from older MDB Studio versions could write "'-1"
        // instead of "-1". Remove that single escape ONLY when its remainder
        // really is a number; never alter a nonnumeric string or a formula.
        if (value.StartsWith('\'') && value.Length > 1)
        {
            string withoutEscape = NormalizeNumericInputNoApostrophe(value[1..]);
            if (decimal.TryParse(withoutEscape, NumberStyles.Float |
                    NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out _) ||
                decimal.TryParse(withoutEscape, NumberStyles.Float |
                    NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out _))
                value = withoutEscape;
        }
        return value;
    }

    private static string NormalizeNumericInputNoApostrophe(string input)
    {
        string value = TrimInvisibleBoundary(input.Trim());
        if (value.Length == 0) return value;
        if (value[0] is '\u2212' or '\uFF0D' or '\u2010' or '\u2011'
            or '\u2012' or '\u2013' or '\uFE63')
            return "-" + value[1..];
        if (value[0] == '\uFF0B') return "+" + value[1..];
        return value;
    }

    private static string TrimInvisibleBoundary(string input) =>
        input.Trim('\uFEFF', '\u200B', '\u2060', '\u200E', '\u200F',
            '\u00A0', '\u202F', '\u0020', '\t', '\r', '\n');

    private static object ParseNumber(string input, Type type)
    {
        string normalized = NormalizeNumericInput(input);
        const NumberStyles integers = NumberStyles.Integer | NumberStyles.AllowThousands;
        const NumberStyles fractions = NumberStyles.Float | NumberStyles.AllowThousands;
        try
        {
            // Use the user's locale for spreadsheets with localized thousands /
            // decimal separators, then retry invariant format for transferred
            // data from other machines. No numeric rounding or Boolean coercion.
            if (type == typeof(int))
                return ParseWithCultures(normalized, (x, c) => int.Parse(x, integers, c));
            if (type == typeof(long))
                return ParseWithCultures(normalized, (x, c) => long.Parse(x, integers, c));
            if (type == typeof(short))
                return ParseWithCultures(normalized, (x, c) => short.Parse(x, integers, c));
            if (type == typeof(byte))
                return ParseWithCultures(normalized, (x, c) => byte.Parse(x, NumberStyles.Integer, c));
            if (type == typeof(double))
                return ParseWithCultures(normalized, (x, c) => double.Parse(x, fractions, c));
            if (type == typeof(decimal))
                return ParseWithCultures(normalized, (x, c) => decimal.Parse(x, fractions, c));
            if (type == typeof(float))
                return ParseWithCultures(normalized, (x, c) => float.Parse(x, fractions, c));
            throw new InvalidOperationException("Unsupported numeric data type.");
        }
        catch (Exception ex) when (ex is FormatException or OverflowException)
        {
            string chars = string.Join(" ", input.EnumerateRunes()
                .Take(40)
                .Select(r => "U+" + r.Value.ToString("X4", CultureInfo.InvariantCulture)));
            // Include code points to make invisible characters and lookalike
            // minus signs diagnosable without requesting the user's MDB data.
            throw new FormatException(
                $"Cannot parse numeric value. Received {input.Length} characters " +
                $"({chars}). Normalized value: '{normalized}'. " +
                "Check for hidden characters, spreadsheet escaping or an invalid number.", ex);
        }
    }

    private static T ParseWithCultures<T>(string input, Func<string, CultureInfo, T> parse)
    {
        try { return parse(input, CultureInfo.CurrentCulture); }
        catch (FormatException) { return parse(input, CultureInfo.InvariantCulture); }
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
