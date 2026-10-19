using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Models.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets built-in and application document properties from an Excel workbook.</summary>
/// <example>
///   <summary>Audit workbook metadata before publishing a report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$properties = Get-OfficeExcelDocumentProperty -Path .\Report.xlsx -Name Title,Company,Department
/// $properties |
///     Format-Table Name, Value, Scope</code>
///   <para>Returns matching built-in, application, and custom workbook properties as structured objects.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelDocumentProperty", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(ExcelDocumentPropertyInfo))]
public sealed class GetOfficeExcelDocumentPropertyCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the workbook.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Property name filter (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Name { get; set; }

    /// <summary>Only return core package properties.</summary>
    [Parameter]
    public SwitchParameter BuiltIn { get; set; }

    /// <summary>Only return application properties such as Company and Manager.</summary>
    [Parameter]
    public SwitchParameter Application { get; set; }

    /// <summary>Only return custom workbook properties.</summary>
    [Parameter]
    public SwitchParameter Custom { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        ExcelDocument? document = null;
        var dispose = false;

        try {
            if (ParameterSetName == ParameterSetPath) {
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                if (!File.Exists(resolvedPath)) {
                    throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
                }

                document = ExcelDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                dispose = true;
            } else {
                document = Document;
            }

            if (document == null) {
                throw new InvalidOperationException("Excel workbook was not provided.");
            }

            var includeBuiltIn = !Application.IsPresent && !Custom.IsPresent || BuiltIn.IsPresent;
            var includeApplication = !BuiltIn.IsPresent && !Custom.IsPresent || Application.IsPresent;
            var includeCustom = !BuiltIn.IsPresent && !Application.IsPresent || Custom.IsPresent;

            IEnumerable<ExcelDocumentPropertyInfo> properties = ExcelDocumentPropertyService.GetProperties(document, includeBuiltIn, includeApplication, includeCustom);
            var patterns = BuildPatterns(Name);
            if (patterns.Count > 0) {
                properties = properties.Where(property => patterns.Any(pattern => pattern.IsMatch(property.Name)));
            }

            WriteObject(properties, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }

    private static List<WildcardPattern> BuildPatterns(string[]? patterns) {
        var compiled = new List<WildcardPattern>();
        foreach (var pattern in patterns ?? Array.Empty<string>()) {
            if (!string.IsNullOrWhiteSpace(pattern)) {
                compiled.Add(new WildcardPattern(pattern, WildcardOptions.IgnoreCase));
            }
        }

        return compiled;
    }
}