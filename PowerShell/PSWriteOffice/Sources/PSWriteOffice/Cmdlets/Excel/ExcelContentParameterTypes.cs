using System;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
namespace PSWriteOffice.Cmdlets.Excel;
/// <summary>Native validation families supported by ExcelValidation.</summary>
public enum ExcelValidationType {
    /// <summary>Restrict entries to an allowed list.</summary>
    List,
    /// <summary>Validate whole-number values.</summary>
    WholeNumber,
    /// <summary>Validate numeric values, including fractions.</summary>
    Decimal,
    /// <summary>Validate dates.</summary>
    Date,
    /// <summary>Validate times.</summary>
    Time,
    /// <summary>Validate the number of text characters.</summary>
    TextLength,
    /// <summary>Validate entries with a custom formula.</summary>
    Formula,
}
/// <summary>Display policy for a single external hyperlink.</summary>
public enum ExcelHyperlinkDisplay {
    /// <summary>Keep the supplied cell text as the label.</summary>
    Text,
    /// <summary>Display the destination URL.</summary>
    Url,
    /// <summary>Display the destination host.</summary>
    Host,
    /// <summary>Use the native smart-link display policy.</summary>
    Smart,
}
/// <summary>Import row projection or ownership-transferring reader output.</summary>
public enum ExcelImportOutput {
    /// <summary>Return each row as a PowerShell object.</summary>
    Object,
    /// <summary>Return each row as a hashtable.</summary>
    Hashtable,
    /// <summary>Return a materialized data table.</summary>
    DataTable,
    /// <summary>Return a native reader that the caller must dispose.</summary>
    DataReader,
    /// <summary>Return the selected column values.</summary>
    Column,
}
/// <summary>Binds Excel dimensions, retaining explicit physical units.</summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class ExcelColumnWidthTransformationAttribute : ArgumentTransformationAttribute {
    /// <summary>Uses points for numeric row heights instead of column character units.</summary>
    public bool RowHeight { get; set; }
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? value = OfficeArgumentValue.Unwrap(inputData);
        if (value is ExcelColumnWidth width) return width;
        if (value is OfficeLength length) return ExcelColumnWidth.Length(length);
        if (value is string text) {
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric))
                return RowHeight ? ExcelColumnWidth.Length(OfficeLength.Points(numeric)) : ExcelColumnWidth.Characters(numeric);
            if (OfficeLength.TryParse(text, out length)) return ExcelColumnWidth.Length(length);
        }
        if (value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal) {
            double numeric = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            return RowHeight ? ExcelColumnWidth.Length(OfficeLength.Points(numeric)) : ExcelColumnWidth.Characters(numeric);
        }
        throw new PSArgumentException("Expected a column character count, row height in points or a physical length such as 2cm.");
    }
}

/// <summary>Comment surfaces returned from a workbook.</summary>
public enum ExcelCommentKind {
    /// <summary>Include notes and threaded comments.</summary>
    All,
    /// <summary>Include legacy cell notes.</summary>
    Note,
    /// <summary>Include threaded comments.</summary>
    Threaded,
}
/// <summary>Native workbook property families.</summary>
public enum ExcelDocumentPropertyKind {
    /// <summary>Select core document properties.</summary>
    BuiltIn,
    /// <summary>Select extended application properties.</summary>
    Application,
    /// <summary>Select named custom properties.</summary>
    Custom,
}
