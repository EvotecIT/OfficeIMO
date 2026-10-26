using System;
using System.Management.Automation;
using OfficeIMO;

namespace PSWriteOffice.Foundation;

/// <summary>Binds the shared format enum by name. Undefined numbers and Unknown are rejected.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeFormatTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? value = OfficeArgumentValue.Unwrap(inputData);
        if (value is OfficeDocumentFormat format && Valid(format)) return format;
        if (value is string text && !int.TryParse(text, out _) && Enum.TryParse(text, true, out format) && Valid(format)) return format;
        throw OfficeArgumentValue.Invalid("Expected a supported OfficeDocumentFormat name.");
    }
    private static bool Valid(OfficeDocumentFormat format) => format != OfficeDocumentFormat.Unknown && Enum.IsDefined(typeof(OfficeDocumentFormat), format);
}
