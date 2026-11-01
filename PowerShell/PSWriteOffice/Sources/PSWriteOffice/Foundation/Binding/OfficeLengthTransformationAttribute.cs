using System;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO;

namespace PSWriteOffice.Foundation;

/// <summary>Binds invariant document lengths; a number means points.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeLengthTransformationAttribute : ArgumentTransformationAttribute {
    /// <summary>Accepts null when the parameter explicitly supports clearing or retaining an override.</summary>
    public bool AllowNull { get; set; }

    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        var value = OfficeArgumentValue.Unwrap(inputData);
        if (value == null && AllowNull) return null!;
        if (value is OfficeLength length) return length;
        if (value is string text && OfficeLength.TryParse(text, out length)) return length;
        if (value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal) {
            string number = ((IFormattable)value).ToString(null, CultureInfo.InvariantCulture);
            if (OfficeLength.TryParse(number, out length)) return length;
        }
        throw OfficeArgumentValue.Invalid("Expected an invariant length such as 12pt, 5cm, 1in, 10mm, 96px or 914400emu; numbers mean points.");
    }
}
