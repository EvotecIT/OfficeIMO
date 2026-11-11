using System.Management.Automation;
using System.Reflection;
using System.Runtime.CompilerServices;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Yaml;

/// <summary>Shared YAML document plumbing for the lifecycle, node and conversion commands.</summary>
internal static class OfficeYamlDocuments {
    // YamlDocument keeps its load path internal and Create() takes none, so the shell remembers the file a live
    // document belongs to. Save-OfficeYaml without -Path and Merge-OfficeDocument read it from here.
    private static readonly ConditionalWeakTable<YamlDocument, string> Paths = new();

    internal static YamlDocument Associate(YamlDocument document, string? path) {
        if (path != null) {
            Paths.Remove(document);
            Paths.Add(document, path);
        }
        return document;
    }

    internal static string? GetPath(YamlDocument document) => Paths.TryGetValue(document, out string? path) ? path : null;

    /// <summary>
    /// Copies caller options so the command switches never change the caller's object. -Untrusted caps every limit at
    /// <see cref="YamlLoadOptions.UntrustedDefaults"/>, so it can only tighten what -Options sets.
    /// </summary>
    internal static YamlLoadOptions BuildLoadOptions(YamlLoadOptions? options, bool untrusted, bool? strict, YamlSchema? schema, int? maxDepth = null) {
        var result = new YamlLoadOptions();
        if (options != null) {
            foreach (PropertyInfo property in typeof(YamlLoadOptions).GetProperties(BindingFlags.Instance | BindingFlags.Public)) {
                if (property.CanRead && property.SetMethod?.IsPublic == true && property.GetIndexParameters().Length == 0)
                    property.SetValue(result, property.GetValue(options));
            }
        }
        if (untrusted) {
            YamlLoadOptions cap = YamlLoadOptions.UntrustedDefaults;
            result.MaxDepth = Math.Min(result.MaxDepth, cap.MaxDepth);
            result.MaxNodes = Math.Min(result.MaxNodes, cap.MaxNodes);
            result.MaxAliasExpansion = Math.Min(result.MaxAliasExpansion, cap.MaxAliasExpansion);
            result.MaxScalarLength = Math.Min(result.MaxScalarLength, cap.MaxScalarLength);
            result.MaxInputBytes = Math.Min(result.MaxInputBytes ?? long.MaxValue, cap.MaxInputBytes!.Value);
        }
        if (strict.HasValue) result.Strict = strict.Value;
        if (schema.HasValue) result.Schema = schema.Value;
        if (maxDepth.HasValue) result.MaxDepth = maxDepth.Value;
        return result;
    }

    /// <summary>
    /// Reports a YAML failure with the engine's stable code as the error ID. The code, line and column come straight
    /// from <see cref="YamlException"/>; when the engine's YamlException-to-OfficeDiagnostic helper lands, this is the
    /// one place to switch over.
    /// </summary>
    internal static ErrorRecord Error(Exception exception, object? input) => exception switch {
        YamlException yaml => OfficeError.Create(yaml, yaml.Code, input, ErrorCategory.InvalidData),
        InvalidDataException => OfficeError.Create(exception, "yaml-multiple-documents", input, ErrorCategory.InvalidData),
        ItemNotFoundException => OfficeError.Create(exception, "yaml-node-not-found", input, ErrorCategory.ObjectNotFound),
        _ => OfficeError.Create(exception, "yaml-failed", input)
    };

    /// <summary>Resolves a path argument to a document, failing clearly when nothing is there.</summary>
    internal static YamlNode Select(YamlDocument document, string path) =>
        document.Select(path) ?? throw new ItemNotFoundException("No YAML node is at '" + path + "'.");

    /// <summary>
    /// Writes a parsed stream as PowerShell data, as ConvertFrom-Json does: a root sequence is enumerated unless
    /// <paramref name="noEnumerate"/>. A stream of several documents needs <paramref name="allDocuments"/>, so a second
    /// document is never dropped silently.
    /// </summary>
    internal static void WriteData(Cmdlet cmdlet, YamlDocument document, bool allDocuments, bool asHashtable, bool noEnumerate) {
        IReadOnlyList<YamlDocumentPart> parts = document.Documents;
        if (parts.Count > 1 && !allDocuments) {
            throw new InvalidDataException("The YAML stream has " + parts.Count + " documents; use -AllDocuments to read all of them.");
        }
        foreach (YamlDocumentPart part in parts) {
            object? value = YamlPowerShellValue.FromYaml(YamlSerializer.Deserialize<object>(part.Root), asHashtable);
            cmdlet.WriteObject(value, enumerateCollection: !noEnumerate);
        }
    }

    /// <summary>Formats a node position as the 1-based line and column PowerShell commands report.</summary>
    internal static PSObject DescribeNode(YamlNode node, string path, bool asHashtable) {
        var result = new PSObject();
        result.Properties.Add(new PSNoteProperty("Path", path));
        result.Properties.Add(new PSNoteProperty("Value", YamlPowerShellValue.FromYaml(YamlSerializer.Deserialize<object>(node), asHashtable)));
        result.Properties.Add(new PSNoteProperty("Node", node));
        result.Properties.Add(new PSNoteProperty("Line", node.Start.Line));
        result.Properties.Add(new PSNoteProperty("Column", node.Start.Column));
        return result;
    }
}
