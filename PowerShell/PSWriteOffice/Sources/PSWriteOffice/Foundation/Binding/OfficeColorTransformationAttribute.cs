using System;
using System.Management.Automation;
using OfficeIMO.Drawing;

namespace PSWriteOffice.Foundation;

/// <summary>Binds Core colors, including CSS names, hex colors and semantic theme roles.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeColorTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) => Bind(inputData);

    internal static OfficeColor Bind(object? inputData) {
        var value = OfficeArgumentValue.Unwrap(inputData);
        if (value is OfficeColor color) return color;
        if (value is string text) {
            if (OfficeColor.TryParse(text, out color) || OfficeColor.TryParseCss(text, out color)) return color;
        }
        throw OfficeArgumentValue.Invalid("Expected an OfficeColor, CSS color name, hex color or theme role.");
    }
}
