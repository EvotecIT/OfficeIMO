using System;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
namespace PSWriteOffice.Cmdlets.Excel;
/// <summary>Native validation families supported by ExcelValidation.</summary>
public enum ExcelValidationType { List, WholeNumber, Decimal, Date, Time, TextLength, Formula }
/// <summary>Display policy for a single external hyperlink.</summary>
public enum ExcelHyperlinkDisplay { Text, Url, Host, Smart }
/// <summary>Import row projection or ownership-transferring reader output.</summary>
public enum ExcelImportOutput { Object, Hashtable, DataTable, DataReader, Column }
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
public enum ExcelCommentKind { All, Note, Threaded }
/// <summary>Native workbook property families.</summary>
public enum ExcelDocumentPropertyKind { BuiltIn, Application, Custom }
