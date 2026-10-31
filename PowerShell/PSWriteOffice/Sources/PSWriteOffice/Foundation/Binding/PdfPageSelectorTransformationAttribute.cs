using System;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Foundation;

/// <summary>Binds native PDF page selectors from expressions and one-based integer positions.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class PdfPageSelectorTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? value = OfficeArgumentValue.Unwrap(inputData);
        if (value is PdfPageSelector selector) return selector;
        if (value is string expression) return PdfPageSelector.Parse(expression);
        if (value is byte or sbyte or short or ushort or int or uint or long or ulong)
            return PdfPageSelector.Parse(Convert.ToString(value, CultureInfo.InvariantCulture)!);
        throw OfficeArgumentValue.Invalid("Expected a native PDF page selector, a range expression, or a one-based integer page.");
    }
}
