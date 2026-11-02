using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Services;

/// <summary>PowerShell binding for existing format families and exact physical extensions, without another public format type.</summary>
internal sealed class OfficeFormatSelection {
    private readonly OfficeDocumentFormat? _family;
    private readonly string? _extension;
    private OfficeFormatSelection(OfficeDocumentFormat? family, string? extension) { _family = family; _extension = extension; }

    /// <summary>Uses the converter's native default only when the caller selected a format family.</summary>
    internal string Extension => _extension ?? OfficeConverter.DefaultExtension(_family!.Value);
    internal OfficeDocumentFormat? Family => _family;
    internal string? PhysicalExtension => _extension;

    internal static OfficeFormatSelection? Parse(string? value) {
        if (value == null) return null;
        if (TryParse(value, out var selection)) return selection;
        throw new PSArgumentException("Expected a supported format name or file extension: " + value);
    }

    internal static bool TryParse(string value, out OfficeFormatSelection? selection) {
        selection = null;
        string text = value.Trim();
        if (!int.TryParse(text, out _) && Enum.TryParse(text, true, out OfficeDocumentFormat family) &&
            family != OfficeDocumentFormat.Unknown && Enum.IsDefined(typeof(OfficeDocumentFormat), family)) {
            selection = new OfficeFormatSelection(family, null);
            return true;
        }
        string extension = text.StartsWith(".", StringComparison.Ordinal) ? text : "." + text;
        bool known = OfficeConversionCapabilityCatalog.All.Any(route =>
            route.SourceExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase) ||
            string.Equals(route.TargetExtension, extension, StringComparison.OrdinalIgnoreCase));
        if (known) selection = new OfficeFormatSelection(null, extension);
        return known;
    }

    internal bool Matches(string extension) {
        if (_extension != null) return string.Equals(extension, _extension, StringComparison.OrdinalIgnoreCase);
        string kind = ReaderCommandUtilities.Reader.DetectKind("document" + extension).ToString();
        return Enum.TryParse(kind, true, out OfficeDocumentFormat family) && family == _family;
    }
}
