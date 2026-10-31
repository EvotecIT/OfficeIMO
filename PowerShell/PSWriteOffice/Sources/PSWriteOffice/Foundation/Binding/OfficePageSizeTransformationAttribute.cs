using System;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO.Pdf;
using OfficeIMO.Drawing;

namespace PSWriteOffice.Foundation;

/// <summary>Binds named paper presets to the existing geometry type used by a format's options.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficePageSizeTransformationAttribute : ArgumentTransformationAttribute {
    /// <summary>Uses the PDF engine's point-based geometry.</summary>
    public OfficePageSizeTransformationAttribute() : this(typeof(PdfPageSize)) { }
    /// <summary>Uses the selected engine's existing page geometry and named presets.</summary>
    public OfficePageSizeTransformationAttribute(Type pageSizeType) {
        PageSizeType = pageSizeType ?? throw new ArgumentNullException(nameof(pageSizeType));
        if (pageSizeType != typeof(PdfPageSize) && pageSizeType != typeof(OfficePageSize))
            throw new ArgumentException("Choose a supported native page geometry type.", nameof(pageSizeType));
    }
    /// <summary>Native geometry type required by the receiving format.</summary>
    public Type PageSizeType { get; }

    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? value = OfficeArgumentValue.Unwrap(inputData);
        if (value == null || PageSizeType.IsInstanceOfType(value)) return value!;
        if (value is string name) {
            if (PageSizeType == typeof(PdfPageSize) && PdfPageSizes.TryGet(name.Trim(), out var size)) return size;
            if (PageSizeType == typeof(OfficePageSize)) {
                var property = typeof(OfficePageSizes).GetProperty(name.Trim(), BindingFlags.Public | BindingFlags.Static | BindingFlags.IgnoreCase);
                if (property?.PropertyType == typeof(OfficePageSize)) return property.GetValue(null)!;
            }
        }
        throw OfficeArgumentValue.Invalid("Expected a known paper size name, such as A4, or the receiving engine's native page dimensions.");
    }
}
