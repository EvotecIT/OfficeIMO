using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData("nim-iwork/simple.numbers")]
    [InlineData("picodocs/sample-v14.4.pages")]
    [InlineData("keynotekit/tabledeck-v15.2.1.key")]
    public void Independently_decoded_empty_extents_and_disabled_filters_do_not_force_visibility_fallback(string relativePath) {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus");
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "empty-hidden-states.json")));
        JsonElement expected = manifest.RootElement.GetProperty("sources").EnumerateArray()
            .Single(item => item.GetProperty("path").GetString() == relativePath);
        string path = Path.Combine(root, relativePath);
        Assert.Equal(expected.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
        IWorkSourceDocument source = IWorkSourceDocument.Open(path);
        IReadOnlyList<IWorkTable> tables;
        IReadOnlyList<IWorkDiagnostic> diagnostics;
        IReadOnlyList<IWorkSourceDeclarationIssue> declarations;
        if (source.Kind == IWorkDocumentKind.Pages) {
            IWorkPagesProjection projection = source.ReadPages();
            tables = projection.Tables; diagnostics = projection.Diagnostics; declarations = projection.SourceDeclarationIssues;
        } else if (source.Kind == IWorkDocumentKind.Numbers) {
            IWorkNumbersProjection projection = source.ReadNumbers();
            tables = projection.Sheets.SelectMany(sheet => sheet.Tables).ToArray();
            diagnostics = projection.Diagnostics; declarations = projection.SourceDeclarationIssues;
        } else {
            IWorkKeynoteProjection projection = source.ReadKeynote();
            tables = projection.Slides.SelectMany(slide => slide.Tables).ToArray();
            diagnostics = projection.Diagnostics; declarations = projection.SourceDeclarationIssues;
        }
        ulong[] modelIds = expected.GetProperty("tables").EnumerateArray()
            .Select(table => table.GetProperty("modelIdentifier").GetUInt64()).OrderBy(id => id).ToArray();
        Assert.Equal(modelIds, tables.Select(table => table.ModelRecord!.Identifier).Distinct().OrderBy(id => id));
        ulong[] filterIds = expected.GetProperty("tables").EnumerateArray()
            .SelectMany(table => table.GetProperty("extents").EnumerateArray())
            .Select(extent => extent.GetProperty("filterIdentifier").GetUInt64()).ToArray();
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_HIDDEN_STATES_UNASSESSED");
        Assert.DoesNotContain(declarations, issue => filterIds.Contains(issue.Owner.RecordIdentifier)
            || modelIds.Contains(issue.Owner.RecordIdentifier) && (issue.FieldPath == "38" || issue.FieldPath.StartsWith("70", StringComparison.Ordinal)));
    }
}
