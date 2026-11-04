using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Drawing;

namespace PSWriteOffice.Foundation;

/// <summary>Binds Core colors, including CSS names, hex colors and semantic theme roles.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeColorTransformationAttribute : ArgumentTransformationAttribute {
    /// <summary>Accepts null when the parameter explicitly supports clearing a color override.</summary>
    public bool AllowNull { get; set; }

    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? value = OfficeArgumentValue.Unwrap(inputData);
        if (AllowNull && value == null) return null!;
        if (value is IEnumerable values && value is not string) {
            var colors = new List<OfficeColor>();
            foreach (object? item in values) colors.Add(Bind(item));
            return colors.ToArray();
        }
        return Bind(value);
    }

    internal static OfficeColor Bind(object? inputData) {
        var value = OfficeArgumentValue.Unwrap(inputData);
        if (value is OfficeColor color) return color;
        if (value is string text) {
            if (OfficeColor.TryParse(text, out color) || OfficeColor.TryParseCss(text, out color)) return color;
        }
        throw OfficeArgumentValue.Invalid("Expected an OfficeColor, CSS color name, hex color or theme role.");
    }
}
