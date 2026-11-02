using System;
using System.IO;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO;

namespace PSWriteOffice.Foundation;

/// <summary>A built-in theme, theme object, or provider-resolved template source.</summary>
/// <remarks>Template extraction belongs to Core's theme reader, not the PowerShell binder.</remarks>
public sealed class OfficeThemeInput {
    private OfficeThemeInput(OfficeTheme? theme, string? templatePath) { Theme = theme; TemplatePath = templatePath; }
    /// <summary>Gets the bound theme, or null when a template must be read.</summary>
    public OfficeTheme? Theme { get; }
    /// <summary>Gets the absolute template path, or null for an already bound theme.</summary>
    public string? TemplatePath { get; }
    /// <summary>Resolves a template through the owning library reader; built-in and supplied objects need no reader.</summary>
    public OfficeTheme Resolve(Func<string, OfficeTheme> readTemplate) => Theme ?? (readTemplate ?? throw new ArgumentNullException(nameof(readTemplate)))(TemplatePath!);
    internal static OfficeThemeInput From(object input, Func<string, string> resolvePath) {
        object? value = OfficeArgumentValue.Unwrap(input);
        if (value is OfficeThemeInput source) return source;
        if (value is OfficeTheme theme) return new(theme, null);
        if (value is string name) {
            foreach (var property in typeof(OfficeTheme).GetProperties(BindingFlags.Public | BindingFlags.Static)) {
                if (property.PropertyType == typeof(OfficeTheme) && property.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    return new((OfficeTheme)property.GetValue(null)!, null);
            }
            string path = resolvePath(name);
            if (!File.Exists(path)) throw OfficeArgumentValue.Invalid("The theme template does not exist: " + path);
            string extension = Path.GetExtension(path).ToLowerInvariant();
            if (extension is not (".potx" or ".pptx" or ".dotx" or ".docx" or ".xltx" or ".xlsx"))
                throw OfficeArgumentValue.Invalid("A theme template must be a Word, Excel or PowerPoint document or template.");
            return new(null, path);
        }
        throw OfficeArgumentValue.Invalid("Expected a built-in theme name, OfficeTheme or template path.");
    }
}

/// <summary>Binds themes while retaining template paths for the owning library reader.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeThemeTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) => OfficeThemeInput.From(inputData,
        path => OfficePath.ResolveOutput(engineIntrinsics.SessionState, path, directory: false));
}
