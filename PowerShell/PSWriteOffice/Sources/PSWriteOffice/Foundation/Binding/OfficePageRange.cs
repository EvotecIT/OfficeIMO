using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Foundation;

/// <summary>Caller-ordered one-based pages using the native page-selection grammar.</summary>
public sealed class OfficePageRange {
    private OfficePageRange(PdfPageSelection selection) { Selection = selection; }
    /// <summary>Gets the immutable native selection, including caller order and repeated pages.</summary>
    public PdfPageSelection Selection { get; }

    /// <summary>Binds native selections, range strings, integers or arrays without a second range parser.</summary>
    public static OfficePageRange From(object value) {
        value = OfficeArgumentValue.Unwrap(value) ?? throw OfficeArgumentValue.Invalid("A page selection is required.");
        if (value is OfficePageRange pages) return pages;
        if (value is PdfPageSelection selection) return new(selection);
        if (value is string ranges) return new(PdfPageSelection.Parse(ranges));
        if (value is Array items) {
            var combined = new List<PdfPageRange>();
            foreach (object item in items) combined.AddRange(From(item).Selection.Ranges);
            return new(PdfPageSelection.FromRanges(combined.ToArray()));
        }
        if (value is byte or sbyte or short or ushort or int or uint or long or ulong) {
            try { return new(PdfPageSelection.From(checked(Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture)))); }
            catch (OverflowException) { throw OfficeArgumentValue.Invalid("Page positions must fit an Int32."); }
        }
        throw OfficeArgumentValue.Invalid("Expected one-based page numbers, a range such as 1-3,5 or a native page selection.");
    }

    /// <inheritdoc />
    public override string ToString() => Selection.ToString();
}

/// <summary>Binds the shared page selection on document image commands.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficePageRangeTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) => OfficePageRange.From(inputData);
}
