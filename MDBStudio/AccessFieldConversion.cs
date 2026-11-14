using System.Globalization;

namespace MDBStudio;

/// <summary>
/// Validates data before an Access field type conversion. Values must be
/// representable without truncation or non-integral numeric rounding.
/// </summary>
public static class AccessFieldConversion
{
    public static object? Normalize(object? raw, AccessFieldDefinition target)
    {
        if (raw is null || raw is DBNull) return null;

        string text = raw is DateTime dt
            ? dt.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture)
            : Convert.ToString(raw, CultureInfo.InvariantCulture) ?? "";

        switch (target.Kind)
        {
            case AccessFieldKind.ShortText:
                if (text.Length > target.Length)
                    throw new FormatException(
                        "Value contains " + text.Length.ToString(CultureInfo.InvariantCulture) +
                        " characters, exceeding the requested text limit of " +
                        target.Length.ToString(CultureInfo.InvariantCulture) + ".");
                return text;
            case AccessFieldKind.LongText:
                return text;
            case AccessFieldKind.Integer:
                if (int.TryParse(text.Trim(), NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out int i)) return i;
                throw new FormatException("Expected a 32-bit whole number.");
            case AccessFieldKind.SmallInteger:
                if (short.TryParse(text.Trim(), NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out short sh)) return sh;
                throw new FormatException("Expected a whole number from -32768 to 32767.");
            case AccessFieldKind.Byte:
                if (byte.TryParse(text.Trim(), NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out byte b)) return b;
                throw new FormatException("Expected a whole number from 0 to 255.");
            case AccessFieldKind.Double:
                if (double.TryParse(text.Trim(), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out double dbl) &&
                    !double.IsInfinity(dbl) && !double.IsNaN(dbl))
                    return dbl;
                throw new FormatException("Expected a finite numeric value.");
            case AccessFieldKind.Currency:
                if (decimal.TryParse(text.Trim(), NumberStyles.Number,
                    CultureInfo.InvariantCulture, out decimal dec) &&
                    dec >= -922337203685477.5808m &&
                    dec <= 922337203685477.5807m &&
                    decimal.Round(dec, 4) == dec)
                    return dec;
                throw new FormatException("Expected a valid currency value with at most four decimal places.");
            case AccessFieldKind.DateTime:
                if (raw is DateTime date) return date;
                if (DateTime.TryParse(text.Trim(), CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces, out var parsedDate))
                    return parsedDate;
                throw new FormatException("Expected a recognizable date and time.");
            case AccessFieldKind.YesNo:
                if (raw is bool flag) return flag;
                if (text.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                    text.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
                    text == "-1" || text == "1") return true;
                if (text.Equals("false", StringComparison.OrdinalIgnoreCase) ||
                    text.Equals("no", StringComparison.OrdinalIgnoreCase) ||
                    text == "0") return false;
                throw new FormatException("Expected Yes, No, True, False, -1, 1 or 0.");
            case AccessFieldKind.Guid:
                if (raw is Guid guid) return guid;
                if (Guid.TryParse(text.Trim(), out var parsedGuid))
                    return parsedGuid;
                throw new FormatException("Expected a GUID.");
            case AccessFieldKind.AutoNumber:
                throw new InvalidOperationException(
                    "Use the complete table designer to change an AutoNumber or primary key.");
            default:
                throw new InvalidOperationException("Unsupported field data type.");
        }
    }

    public static bool Matches(object? left, object? right,
        AccessFieldDefinition target)
    {
        object? a = Normalize(left, target);
        object? b = Normalize(right, target);
        if (a is null || b is null) return a is null && b is null;
        if (a is double da && b is double db)
            return Math.Abs(da - db) <= Math.Max(1e-10, Math.Max(Math.Abs(da), Math.Abs(db)) * 1e-12);
        return Equals(a, b);
    }

    public static string DisplaySample(object? value)
    {
        if (value is null || value is DBNull) return "NULL";
        string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        if (text.Length > 90) text = text[..90] + "...";
        return "'" + text.Replace("\r", " ").Replace("\n", " ") + "'";
    }
}
