using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO.GoogleWorkspace;

namespace PSWriteOffice.Services.Google;

internal static class GoogleWorkspaceScopes {
    internal static readonly IReadOnlyDictionary<string, string> Names = typeof(GoogleWorkspaceScopeCatalog)
        .GetFields(BindingFlags.Public | BindingFlags.Static).Where(field => field.IsLiteral && field.FieldType == typeof(string))
        .ToDictionary(field => field.Name, field => (string)field.GetRawConstantValue()!, StringComparer.OrdinalIgnoreCase);
    internal static string[] Resolve(string[]? values) => (values ?? new[] {
        GoogleWorkspaceScopeCatalog.DriveFile, GoogleWorkspaceScopeCatalog.Documents, GoogleWorkspaceScopeCatalog.Spreadsheets
    }).Select(value => Names.TryGetValue(value, out var scope) ? scope :
        Uri.TryCreate(value, UriKind.Absolute, out var uri) && uri.Scheme == Uri.UriSchemeHttps ? value :
        throw new PSArgumentException("Use a native Google Workspace scope name or its full HTTPS URL.")).Distinct(StringComparer.Ordinal).ToArray();
}

/// <summary>Completes scope names from the native Google Workspace scope catalog.</summary>
public sealed class OfficeGoogleWorkspaceScopeCompleter : IArgumentCompleter {
    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(string commandName, string parameterName, string wordToComplete,
        System.Management.Automation.Language.CommandAst commandAst, IDictionary fakeBoundParameters) {
        foreach (var name in GoogleWorkspaceScopes.Names.Keys.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))
            if (name.StartsWith(wordToComplete ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                yield return new CompletionResult(name, name, CompletionResultType.ParameterValue, GoogleWorkspaceScopes.Names[name]);
    }
}
