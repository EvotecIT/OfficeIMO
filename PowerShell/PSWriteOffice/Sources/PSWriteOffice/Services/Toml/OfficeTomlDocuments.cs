using System.Management.Automation;
using System.Text;
using OfficeIMO.Toml;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Toml;

/// <summary>Shared reading, writing and error mapping of the TOML conversion commands.</summary>
internal static class OfficeTomlDocuments {
    internal static TomlLoadOptions LoadOptions(bool untrusted, int? depth) {
        TomlLoadOptions options = untrusted ? TomlLoadOptions.UntrustedDefaults : new TomlLoadOptions();
        if (depth is int max) options.MaxDepth = max;
        return options;
    }

    /// <summary>Writes the document as one object: its root table.</summary>
    internal static void WriteData(Cmdlet cmdlet, TomlDocument document, bool asHashtable) =>
        cmdlet.WriteObject(TomlPowerShellValue.FromToml(document.Root, asHashtable));

    /// <summary>Builds TOML text from collected input and warns about what TOML cannot hold.</summary>
    internal static string Serialize(Cmdlet cmdlet, object? value, int depth, string rootName) {
        using TomlDocument document = TomlPowerShellValue.ToDocument(value, depth, rootName, out bool wrapped, out int nulls, out bool truncated);
        if (truncated) cmdlet.WriteWarning("Resulting TOML is truncated as serialization has exceeded the set depth of " + depth + ".");
        if (wrapped) cmdlet.WriteWarning("TOML documents are tables, so the input was written under the key '" + rootName + "'.");
        if (nulls > 0) cmdlet.WriteWarning("TOML has no null, so " + nulls + " null value" + (nulls == 1 ? " was" : "s were") + " omitted.");
        return document.ToString();
    }

    internal static ErrorRecord Error(Exception exception, object? input) => exception switch {
        TomlException toml => OfficeError.Create(exception, "toml-" + toml.Code, input, ErrorCategory.InvalidData),
        _ => OfficeError.Create(exception, "toml-failed", input)
    };
}

/// <summary>
/// Collects ConvertTo-/Export-OfficeToml input the way ConvertTo-Json does: one piped object is written as itself,
/// several piped objects as one array, and an -InputObject argument exactly as given.
/// </summary>
internal sealed class OfficeTomlWriterInput {
    private readonly List<object?> _items = new();
    private bool _fromArgument;

    internal void Add(PSCmdlet cmdlet, object? value) {
        _fromArgument = !cmdlet.MyInvocation.ExpectingInput;
        _items.Add(value);
    }

    internal object? Value => _fromArgument || _items.Count == 1 ? _items[_items.Count - 1] : _items.ToArray();

    internal bool HasInput => _items.Count > 0;
}
