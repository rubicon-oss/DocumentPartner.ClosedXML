using System;
using System.Globalization;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Converts text of a mapped XML element or attribute into a cell value according to the
    /// <c>xmlDataType</c> of the mapping on import. The rules are
    /// specified by the tests in <c>ClosedXML.XmlMaps.Tests</c>.
    /// </summary>
    internal static class XmlDataTypeConverter
    {
        /// <summary>
        /// Maximum length of text in a cell.
        /// </summary>
        internal const int MaxTextLength = 32767;

        internal static readonly char[] XmlWhitespace = { ' ', '\t', '\r', '\n' };

        private static readonly DateTime Date1904 = new(1904, 1, 1);
        private static readonly DateTime Date1900Base = new(1899, 12, 31);
        private static readonly DateTime Date1900LeapBug = new(1900, 3, 1);

        /// <summary>
        /// Convert the text of an XML node to a cell value.
        /// </summary>
        /// <param name="text">The text of the element or attribute, with the leading and trailing whitespace already removed.</param>
        /// <param name="xmlDataType">The <c>xmlDataType</c> of the mapping, e.g. <c>int</c> or <c>date</c>.</param>
        /// <param name="use1904DateSystem">Whether the workbook uses the 1904 date system.</param>
        /// <param name="truncated">Set to <c>true</c>, if the text was too long for a cell and was truncated.</param>
        internal static XLCellValue Convert(string text, string? xmlDataType, bool use1904DateSystem, out bool truncated)
        {
            truncated = false;
            if (text.Length == 0)
                return Blank.Value;

            if (TryConvert(text, xmlDataType, use1904DateSystem, out var value))
                return value;

            if (text.Length > MaxTextLength)
            {
                truncated = true;
                return text.Substring(0, MaxTextLength);
            }

            return text;
        }

        /// <summary>
        /// Built-in number format a cell gets after an import. Usually only a cell with the General format gets the format
        /// of the XML data type. For dates, built-in formats that don't fit the value are changed too: a value of 0 (the 1st
        /// of January 1904 or --01-01 in the 1904 date system) gets a time format, a value of at least one day in a time format
        /// (24:00:00) gets a date format. Custom formats are kept.
        /// </summary>
        /// <param name="xmlDataType">The <c>xmlDataType</c> of the mapping.</param>
        /// <param name="value">The imported value.</param>
        /// <param name="currentFormatId">The built-in number format of the cell, or <c>null</c> for a custom format.</param>
        /// <returns>The new number format id or <c>null</c>, if the format of the cell is kept.</returns>
        internal static int? GetNumberFormatId(string? xmlDataType, XLCellValue value, int? currentFormatId)
        {
            if (value.IsBlank || currentFormatId is not { } formatId)
                return null;

            var isGeneral = formatId == (int)XLPredefinedFormat.General;
            var isDate = xmlDataType is "date" or "dateTime" or "time" or "gYearMonth" or "gMonthDay";
            if (isDate && value.IsUnifiedNumber)
            {
                var number = value.GetUnifiedNumber();
                if (number == 0 && (isGeneral || formatId is 14 or 15 or 16 or 17 or 22))
                    return (int)XLPredefinedFormat.DateTime.Hour24MinutesSeconds;

                if (number >= 1 && (formatId is >= 18 and <= 21 || (isGeneral && xmlDataType == "time")))
                    return number == Math.Floor(number)
                        ? (int)XLPredefinedFormat.DateTime.DayMonthYear4WithSlashes
                        : (int)XLPredefinedFormat.DateTime.MonthDayYear4WithDashesHour24Minutes;
            }

            return isGeneral ? GetDefaultNumberFormatId(xmlDataType) : null;
        }

        /// <summary>
        /// Built-in number format for cells of an XML data type, when the cell has the General format.
        /// </summary>
        /// <returns>The number format id or <c>null</c>, if the General format is kept.</returns>
        internal static int? GetDefaultNumberFormatId(string? xmlDataType)
        {
            switch (xmlDataType)
            {
                case null:
                case "byte":
                case "unsignedByte":
                case "short":
                case "unsignedShort":
                case "int":
                case "unsignedInt":
                case "long":
                case "unsignedLong":
                case "integer":
                case "positiveInteger":
                case "negativeInteger":
                case "nonPositiveInteger":
                case "nonNegativeInteger":
                case "decimal":
                case "float":
                case "double":
                case "boolean":
                    return null;
                case "date":
                    return (int)XLPredefinedFormat.DateTime.DayMonthYear4WithSlashes; // 14, mm-dd-yy
                case "dateTime":
                    return (int)XLPredefinedFormat.DateTime.MonthDayYear4WithDashesHour24Minutes; // 22, m/d/yy h:mm
                case "time":
                    return (int)XLPredefinedFormat.DateTime.Hour24MinutesSeconds; // 21, h:mm:ss
                case "gYear":
                case "gMonth":
                case "gDay":
                    return (int)XLPredefinedFormat.Number.Integer; // 1, 0
                case "gYearMonth":
                    return (int)XLPredefinedFormat.DateTime.MonthAbbrYear2WithDash; // 17, mmm-yy
                case "gMonthDay":
                    return (int)XLPredefinedFormat.DateTime.DayMonthAbbrWithDash; // 16, d-mmm
                default:
                    return (int)XLPredefinedFormat.Number.Text; // 49, @
            }
        }

        /// <summary>
        /// Convert a cell value to the text of an exported XML element or attribute.
        /// </summary>
        /// <returns>The text, an empty text for errors, or <c>null</c> for a blank cell.</returns>
        internal static string? ToXmlText(XLCellValue value, string? xmlDataType, bool use1904DateSystem)
        {
            switch (value.Type)
            {
                case XLDataType.Blank:
                    return null;
                case XLDataType.Text:
                    // Line breaks are written as CR LF.
                    return value.GetText().Replace("\r\n", "\n").Replace("\n", "\r\n");
                case XLDataType.Boolean:
                    return value.GetBoolean() ? "true" : "false";
                case XLDataType.Error:
                    return string.Empty;
            }

            var number = GetSerial(value, use1904DateSystem);
            switch (xmlDataType)
            {
                case "date":
                    return TrySplitSerial(number, use1904DateSystem, out var date, out _) ? date : FormatNumber(number);
                case "dateTime":
                    return TrySplitSerial(number, use1904DateSystem, out var day, out var milliseconds)
                        ? day + "T" + FormatTime(milliseconds)
                        : FormatNumber(number);
                case "time":
                    return TrySplitSerial(number, use1904DateSystem, out _, out var timeOfDay) ? FormatTime(timeOfDay) : FormatNumber(number);
                case "gYearMonth":
                    return TrySplitSerial(number, use1904DateSystem, out var yearMonth, out _) ? yearMonth.Substring(0, 7) : FormatNumber(number);
                case "gMonthDay":
                    return TrySplitSerial(number, use1904DateSystem, out var monthDay, out _) ? "--" + monthDay.Substring(5) : FormatNumber(number);
                case "gYear":
                    // Other numbers are written with 17 digits, like the other g types.
                    return number != 0 && number == Math.Floor(number) && Math.Abs(number) <= int.MaxValue
                        ? (number < 0 ? "-" : string.Empty) + Math.Abs(number).ToString("0000", CultureInfo.InvariantCulture)
                        : FormatRoundTrip(number);
                case "gDay":
                    return number >= 1 && number <= 31 && number == Math.Floor(number)
                        ? "---" + number.ToString("00", CultureInfo.InvariantCulture)
                        : FormatRoundTrip(number);
                case "gMonth":
                    // The notation of the first version of XML Schema is written.
                    return number >= 1 && number <= 12 && number == Math.Floor(number)
                        ? "--" + number.ToString("00", CultureInfo.InvariantCulture) + "--"
                        : FormatRoundTrip(number);
                default:
                    return FormatNumber(number, xmlDataType);
            }
        }

        /// <summary>
        /// Format a number on export: rounded to 15 significant digits. Numbers smaller than the smallest number of a cell
        /// are 0. The digits are written without exponent if they fit into 21 characters (without sign), otherwise with
        /// exponent (1E+21, 1E-20). Numbers of decimal types (decimal, integer, int, ...) are always written without exponent.
        /// </summary>
        internal static string FormatNumber(double number, string? xmlDataType = null)
        {
            if (Math.Abs(number) < 2.2251E-308)
                return "0";

            var text = number.ToString("G15", CultureInfo.InvariantCulture);
            var exponentIndex = text.IndexOf('E');
            if (exponentIndex < 0)
                return text;

            var exponent = int.Parse(text.Substring(exponentIndex + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
            var mantissa = text.Substring(0, exponentIndex);
            var negative = mantissa.StartsWith("-", StringComparison.Ordinal);
            var digits = mantissa.TrimStart('-').Replace(".", string.Empty);
            var pointPosition = 1 + exponent;
            string result;
            if (pointPosition <= 0)
                result = "0." + new string('0', -pointPosition) + digits;
            else if (pointPosition >= digits.Length)
                result = digits + new string('0', pointPosition - digits.Length);
            else
                result = digits.Substring(0, pointPosition) + "." + digits.Substring(pointPosition);

            // The exponent has at least two digits like in G15 (1E-07).
            if (result.Length > 21 && !IsDecimalType(xmlDataType))
                return text;

            return negative ? "-" + result : result;
        }

        private static bool IsDecimalType(string? xmlDataType) => xmlDataType is "decimal" or "integer" or "nonPositiveInteger" or
            "negativeInteger" or "long" or "int" or "short" or "byte" or "nonNegativeInteger" or "unsignedLong" or "unsignedInt" or
            "unsignedShort" or "unsignedByte" or "positiveInteger";

        /// <summary>
        /// Format a number with 17 significant digits, used for numbers that don't fit <c>gYear</c>, <c>gMonth</c> or <c>gDay</c>.
        /// </summary>
        private static string FormatRoundTrip(double number)
        {
            var text = number.ToString("G17", CultureInfo.InvariantCulture);
            return text == "-0" ? "0" : text;
        }

        private static string FormatTime(long milliseconds) =>
            TimeSpan.FromMilliseconds(milliseconds).ToString("hh':'mm':'ss'.'fff", CultureInfo.InvariantCulture);

        /// <summary>
        /// Get the serial date of a number in the date system of the workbook. ClosedXML gives the serial date of the 1900 date system for a date.
        /// </summary>
        private static double GetSerial(XLCellValue value, bool use1904DateSystem)
        {
            if (value.IsDateTime && use1904DateSystem)
                return (value.GetDateTime() - Date1904).TotalDays;

            return value.GetUnifiedNumber();
        }

        /// <summary>
        /// Split a serial date into the date text and the milliseconds of the day on export. The number is rounded to
        /// milliseconds first, so a time just before midnight belongs to the next day.
        /// <list type="bullet">
        /// <item>Dates are exported from the 31st of December 1499 (1900 date system) or the 1st of January 1504 (1904 date system),
        /// the range is checked before rounding.</item>
        /// <item>In the 1900 date system the 29th of February 1900 is counted and written. The serial dates skip
        /// 10 days before the 2nd of March 1500.</item>
        /// <item>In the 1904 date system a negative serial date is a day later, so -1 and -0.5 are both the 1st of January 1904.</item>
        /// </list>
        /// </summary>
        /// <returns><c>false</c>, if the number isn't in the range of exported dates.</returns>
        private static bool TrySplitSerial(double serial, bool use1904DateSystem, out string date, out long milliseconds)
        {
            date = string.Empty;
            milliseconds = 0;
            if (!(serial >= (use1904DateSystem ? -146098 : -146087) && serial < (use1904DateSystem ? 2957004 : 2958466)))
                return false;

            var total = (long)Math.Round(serial * 86_400_000d);
            var day = (long)Math.Floor(total / 86_400_000d);
            milliseconds = total - day * 86_400_000;
            if (day > (use1904DateSystem ? 2957003 : 2958465))
                return false;

            if (use1904DateSystem)
            {
                date = FormatDate(Date1904.AddDays(serial < 0 ? day + 1 : day));
                return true;
            }

            if (day == 60)
            {
                date = "1900-02-29";
                return true;
            }

            var gregorian = day < 60 ? Date1900Base.AddDays(day) : Date1900Base.AddDays(day - 1);
            date = FormatDate(day <= -146037 ? gregorian.AddDays(-10) : gregorian);
            return true;
        }

        private static string FormatDate(DateTime date) => date.ToString("yyyy'-'MM'-'dd", CultureInfo.InvariantCulture);

        private static bool TryConvert(string text, string? xmlDataType, bool use1904DateSystem, out XLCellValue value)
        {
            value = Blank.Value;
            switch (xmlDataType)
            {
                case "byte":
                case "unsignedByte":
                case "short":
                case "unsignedShort":
                case "int":
                case "unsignedInt":
                case "long":
                case "unsignedLong":
                case "integer":
                case "positiveInteger":
                case "negativeInteger":
                case "nonPositiveInteger":
                case "nonNegativeInteger":
                case "decimal":
                case "float":
                case "double":
                    if (!TryParseNumber(text, out var number))
                        return false;

                    value = number;
                    return true;

                case "gYear":
                    // At least four digits like xs:gYear, but any digits after a plus sign are taken too.
                    if (!IsYear(text) || !TryParseNumber(text, out var year))
                        return false;

                    value = year;
                    return true;

                case "boolean":
                    // Only the literals true and false become logical values, other numbers (1, 0, -1, 1.5) stay numbers.
                    if (text == "true" || text == "false")
                    {
                        value = text == "true";
                        return true;
                    }

                    if (!TryParseNumber(text, out var booleanNumber))
                        return false;

                    value = booleanNumber;
                    return true;

                case "date":
                    return TryParseDate(text, out var date) && TryGetDateValue(date, use1904DateSystem, out value);

                case "dateTime":
                    var separator = text.IndexOf('T');
                    return separator > 0 &&
                           TryParseDate(text.Substring(0, separator), out var day) &&
                           TryParseTime(text.Substring(separator + 1), out var timeOfDay) &&
                           (timeOfDay < TimeSpan.FromDays(1) || day < DateTime.MaxValue.Date) &&
                           TryGetDateValue(day + timeOfDay, use1904DateSystem, out value);

                case "time":
                    if (!TryParseTime(text, out var time))
                        return false;

                    // 24:00:00 is a valid xs:time, it is stored as a whole day.
                    value = time == TimeSpan.FromDays(1) ? 1 : time;
                    return true;

                case "gYearMonth":
                    return text.Length == 7 && text[4] == '-' &&
                           TryParseDate(text + "-01", out var firstOfMonth) &&
                           TryGetDateValue(firstOfMonth, use1904DateSystem, out value);

                case "gMonthDay":
                    // The day is put into the year 1900 with a February of 29 days and isn't checked against the month,
                    // so --02-31 is the 2nd of March. In the 1904 date system it's the same number minus one.
                    if (text.Length != 7 || !text.StartsWith("--", StringComparison.Ordinal) || text[4] != '-' ||
                        !TryParseTwoDigits(text, 2, 1, 12, out var month) || !TryParseTwoDigits(text, 5, 1, 31, out var dayInMonth))
                        return false;

                    var dayOfYear = DaysBeforeMonth1900[month - 1] + dayInMonth;
                    value = use1904DateSystem ? dayOfYear - 1 : dayOfYear;
                    return true;

                case "gMonth":
                    // Only the notation of the first version of XML Schema, --05-- but not --05.
                    if (text.Length != 6 || !text.StartsWith("--", StringComparison.Ordinal) || !text.EndsWith("--", StringComparison.Ordinal) ||
                        !TryParseTwoDigits(text, 2, 1, 12, out var monthOfYear))
                        return false;

                    value = monthOfYear;
                    return true;

                case "gDay":
                    if (text.Length != 5 || !text.StartsWith("---", StringComparison.Ordinal) ||
                        !TryParseTwoDigits(text, 3, 1, 31, out var dayOfMonth))
                        return false;

                    value = dayOfMonth;
                    return true;
            }

            return false;
        }

        private static readonly int[] DaysBeforeMonth1900 = { 0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335 };

        private static bool TryParseTwoDigits(string text, int index, int min, int max, out int number)
        {
            return int.TryParse(text.Substring(index, 2), NumberStyles.None, CultureInfo.InvariantCulture, out number) &&
                   number >= min && number <= max;
        }

        /// <summary>
        /// Check the spelling of a <c>gYear</c> that is converted to a number: an optional minus sign and at least four digits, not 0000,
        /// or a plus sign with any digits.
        /// </summary>
        private static bool IsYear(string text)
        {
            var start = text.Length > 0 && (text[0] == '-' || text[0] == '+') ? 1 : 0;
            var digits = text.Length - start;
            for (var i = start; i < text.Length; i++)
            {
                if (text[i] < '0' || text[i] > '9')
                    return false;
            }

            if (digits == 0)
                return false;

            return text[0] == '+' || (digits >= 4 && text.Substring(start).Trim('0').Length > 0);
        }

        /// <summary>
        /// Parse a number. Whitespace isn't allowed, whitespace of a CDATA section is kept and such a number isn't converted.
        /// A positive number too small for a double (1E-400) or with an exponent out of range (0.0E-999) stays text, like a number
        /// too large. A negative one (-1E-400) is taken as 0.
        /// </summary>
        private static bool TryParseNumber(string text, out double number)
        {
            const NumberStyles style = NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent;
            if (!double.TryParse(text, style, CultureInfo.InvariantCulture, out number) || double.IsNaN(number) || double.IsInfinity(number))
                return false;

            if (number != 0 || text.StartsWith("-", StringComparison.Ordinal))
                return true;

            var exponentIndex = text.IndexOfAny(new[] { 'e', 'E' });
            var mantissa = exponentIndex < 0 ? text : text.Substring(0, exponentIndex);
            if (mantissa.IndexOfAny("123456789".ToCharArray()) >= 0)
                return false;

            return exponentIndex < 0 ||
                   !int.TryParse(text.Substring(exponentIndex + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var exponent) ||
                   exponent >= -324;
        }

        /// <summary>
        /// Parse <c>yyyy-MM-dd</c> without time zone. Dates with a time zone are kept as text.
        /// </summary>
        private static bool TryParseDate(string text, out DateTime date)
        {
            return DateTime.TryParseExact(text, "yyyy'-'MM'-'dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
        }

        /// <summary>
        /// Parse <c>HH:mm:ss</c> with an optional fraction of zeros, or <c>HH:mm</c>. Times with a non-zero fraction
        /// of a second or with a time zone are kept as text.
        /// </summary>
        private static bool TryParseTime(string text, out TimeSpan time)
        {
            time = TimeSpan.Zero;
            if (text.Length == 5)
                text += ":00";

            if (text.Length < 8 || text[2] != ':' || text[5] != ':')
                return false;

            if (text.Length > 8)
            {
                if (text[8] != '.' || text.Length == 9)
                    return false;

                for (var i = 9; i < text.Length; i++)
                {
                    if (text[i] != '0')
                        return false;
                }
            }

            if (!int.TryParse(text.Substring(0, 2), NumberStyles.None, CultureInfo.InvariantCulture, out var hours) ||
                !int.TryParse(text.Substring(3, 2), NumberStyles.None, CultureInfo.InvariantCulture, out var minutes) ||
                !int.TryParse(text.Substring(6, 2), NumberStyles.None, CultureInfo.InvariantCulture, out var seconds))
                return false;

            var isEndOfDay = hours == 24 && minutes == 0 && seconds == 0;
            if ((hours > 23 && !isEndOfDay) || minutes > 59 || seconds > 59)
                return false;

            time = new TimeSpan(hours, minutes, seconds);
            return true;
        }

        /// <summary>
        /// Get the cell value of a date. Dates ClosedXML can't represent as <see cref="DateTime"/> serial
        /// (the days before 1900-03-01, affected by the 29th of February 1900 of the 1900 date system) are stored as numbers, as are all dates of
        /// the 1904 date system, because ClosedXML saves the first days of 1904 wrongly.
        /// </summary>
        private static bool TryGetDateValue(DateTime date, bool use1904DateSystem, out XLCellValue value)
        {
            value = Blank.Value;
            if (use1904DateSystem)
            {
                if (date < Date1904)
                    return false;

                value = (date - Date1904).TotalDays;
                return true;
            }

            if (date <= Date1900Base)
                return false;

            value = date >= Date1900LeapBug ? date : GetSerial1900(date);
            return true;
        }

        private static double GetSerial1900(DateTime date)
        {
            var serial = (date - Date1900Base).TotalDays;
            return date >= Date1900LeapBug ? serial + 1 : serial;
        }
    }
}
