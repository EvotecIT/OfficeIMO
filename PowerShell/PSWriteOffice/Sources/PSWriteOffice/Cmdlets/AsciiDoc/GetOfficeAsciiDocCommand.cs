using System.IO;
using System.Management.Automation;
using OfficeIMO.AsciiDoc;

namespace PSWriteOffice.Cmdlets.AsciiDoc;

/// <summary>Parses an AsciiDoc file or source string into OfficeIMO's native document model.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeAsciiDoc", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(AsciiDocDocument), typeof(AsciiDocParseResult))]
public sealed class GetOfficeAsciiDocCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";

    /// <summary>Path to an AsciiDoc file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>AsciiDoc source text.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Optional parser settings.</summary>
    [Parameter]
    public AsciiDocParseOptions? Options { get; set; }

    /// <summary>Return the parse result with diagnostics instead of only the document.</summary>
    [Parameter]
    public SwitchParameter AsResult { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        AsciiDocParseResult result;
        if (ParameterSetName == ParameterSetPath)
        {
            var path = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!File.Exists(path)) throw new FileNotFoundException($"File '{path}' was not found.", path);
            result = AsciiDocDocument.LoadResult(path, Options);
        }
        else
        {
            result = AsciiDocDocument.ParseResult(Text ?? string.Empty, Options);
        }

        WriteObject(AsResult.IsPresent ? result : result.Document);
    }
}
