using System.Management.Automation;
using OfficeIMO.Yaml;

namespace PSWriteOffice.Services.Yaml;

/// <summary>
/// Collects ConvertTo-/Export-OfficeYaml input the way ConvertTo-Json does: one piped object is written as itself,
/// several piped objects as one sequence, and an -InputObject argument exactly as given.
/// </summary>
internal sealed class OfficeYamlWriterInput {
    private readonly List<object?> _items = new();
    private bool _fromArgument;

    internal void Add(PSCmdlet cmdlet, object? value) {
        _fromArgument = !cmdlet.MyInvocation.ExpectingInput;
        _items.Add(value);
    }

    /// <summary>Projects the collected input and serializes it; warns once when Depth truncated a value.</summary>
    internal string Serialize(Cmdlet cmdlet, int depth, YamlWriterOptions? writerOptions) {
        object? value = _fromArgument || _items.Count == 1 ? _items[_items.Count - 1] : _items.ToArray();
        bool truncated = false;
        object? projected = YamlPowerShellValue.ToYaml(value, depth, ref truncated);
        if (truncated) cmdlet.WriteWarning("Resulting YAML is truncated as serialization has exceeded the set depth of " + depth + ".");
        var options = new YamlSerializerOptions { WriterOptions = writerOptions };
        return YamlSerializer.Serialize(projected, projected?.GetType() ?? typeof(object), options);
    }

    internal bool HasInput => _items.Count > 0;
}
