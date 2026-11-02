using System;
using System.Globalization;
using System.Management.Automation;
using System.Text;

namespace PSWriteOffice.Foundation;

/// <summary>Binds encoding objects, standard PowerShell encoding names and available code pages.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeEncodingTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? value = OfficeArgumentValue.Unwrap(inputData);
        if (value == null) return null!;
        if (value is Encoding encoding) return encoding.Clone();
        try {
            if (value is byte or short or int or long) return Encoding.GetEncoding(Convert.ToInt32(value, CultureInfo.InvariantCulture));
            if (value is string name) return name.Trim().ToLowerInvariant() switch {
                "utf8" or "utf8nobom" => new UTF8Encoding(false),
                "utf8bom" => new UTF8Encoding(true),
                "unicode" or "utf16" => Encoding.Unicode.Clone(),
                "bigendianunicode" or "utf16be" => Encoding.BigEndianUnicode.Clone(),
                "utf32" => Encoding.UTF32.Clone(),
                "ascii" => Encoding.ASCII.Clone(),
                "default" => Encoding.Default.Clone(),
                "latin1" => Encoding.GetEncoding("iso-8859-1"),
                _ => Encoding.GetEncoding(name)
            };
        } catch (Exception error) when (error is ArgumentException or NotSupportedException or OverflowException) {
            throw OfficeArgumentValue.Invalid("Unknown or unavailable text encoding: " + value + ". " + error.Message);
        }
        throw OfficeArgumentValue.Invalid("Expected an Encoding instance, encoding name or numeric code page.");
    }
}
