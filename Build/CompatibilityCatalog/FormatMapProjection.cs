using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO;
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

/// <summary>
/// Projects the conversion catalog into the website's format map: every route with the surfaces that can run it.
/// Surface membership comes from the code each surface runs, not from a hand-kept list:
/// .NET exposes every catalog route, the browser engine marks its own routes, Studio lists the iWork-enabled
/// workflow runner's routes, and the CLI runs the workflow runner (workflow batch --route) and the iWork runner (convert).
/// </summary>
internal static class FormatMapProjection {
    private const string PowerShellRoutesFile = "powershell-routes.json";
    private const string PowerShellSnapshot = "Website/data/apidocs/powershell/command-metadata.json";

    internal static string Create(string repositoryRoot) {
        HashSet<string> studio = Ids(IWorkWorkflow.CreateRunner().ConversionRoutes);
        HashSet<string> cli = Ids(new OfficeWorkflowRunner().ConversionRoutes);
        cli.UnionWith(studio);
        Dictionary<string, PowerShellRoute> powershell = LoadPowerShellRoutes(repositoryRoot);

        FormatMapRoute[] routes = OfficeConversionCapabilityCatalog.All
            .OrderBy(static route => route.Id, StringComparer.Ordinal)
            .Select(route => {
                powershell.TryGetValue(route.Id, out PowerShellRoute? shell);
                return new FormatMapRoute(
                    route.Id,
                    route.Source,
                    route.Target,
                    route.PackageId,
                    route.SupportLevel.ToString(),
                    route.Fidelity.ToString(),
                    route.Api,
                    Surfaces(route, studio, cli, shell is not null),
                    shell?.Cmdlet,
                    shell?.Example);
            })
            .ToArray();
        string[] unknown = powershell.Keys.Except(routes.Select(static route => route.Id), StringComparer.Ordinal).ToArray();
        if (unknown.Length > 0) {
            throw new InvalidOperationException($"{PowerShellRoutesFile} names routes the conversion catalog doesn't have: {string.Join(", ", unknown)}");
        }
        FormatMapFormat[] formats = routes
            .SelectMany(static route => new[] { route.Source, route.Target })
            .Distinct(StringComparer.Ordinal)
            .Select(static name => new FormatMapFormat(name, Family(name)))
            .OrderBy(static format => Array.IndexOf(Families, format.Family))
            .ThenBy(static format => Rank(format.Name))
            .ThenBy(static format => format.Name, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        // The website embeds these strings in a JSON script block without escaping; keep them plain.
        foreach (FormatMapRoute route in routes) {
            foreach (string value in new[] { route.Id, route.Source, route.Target, route.Package, route.Api, route.Cmdlet ?? "", route.Example ?? "" }) {
                if (value.Any(static ch => ch is '"' or '\\' or '<' or '>' or '&' || char.IsControl(ch))) {
                    throw new InvalidOperationException($"Route '{route.Id}' has a value the format map can't embed as-is: {value}");
                }
            }
        }
        var model = new FormatMapModel(1, Families, formats, routes);
        return JsonSerializer.Serialize(model, FormatMapJsonContext.Default.FormatMapModel).Replace("\r\n", "\n") + "\n";
    }

    private static HashSet<string> Ids(IEnumerable<OfficeWorkflowRoute> routes) =>
        new(routes.Select(static route => route.Id), StringComparer.Ordinal);

    private static string[] Surfaces(OfficeConversionCapability route, HashSet<string> studio, HashSet<string> cli, bool powershell) {
        var surfaces = new List<string> { "dotnet" };
        if (route.BrowserAvailable) surfaces.Add("browser");
        if (cli.Contains(route.Id)) surfaces.Add("cli");
        if (studio.Contains(route.Id)) surfaces.Add("studio");
        if (powershell) surfaces.Add("powershell");
        return surfaces.ToArray();
    }

    /// <summary>
    /// PSWriteOffice routes: a reviewed route-to-cmdlet list kept beside this tool. Each cmdlet must exist in the PowerShell
    /// command snapshot the website publishes, so a renamed or removed cmdlet fails generation instead of shipping a dead claim.
    /// </summary>
    private static Dictionary<string, PowerShellRoute> LoadPowerShellRoutes(string repositoryRoot) {
        string listPath = Path.Combine(repositoryRoot, "Build", "CompatibilityCatalog", PowerShellRoutesFile);
        string snapshotPath = Path.Combine(repositoryRoot, PowerShellSnapshot.Replace('/', Path.DirectorySeparatorChar));
        PowerShellRoute[] listed = JsonSerializer.Deserialize(File.ReadAllText(listPath), FormatMapJsonContext.Default.PowerShellRouteArray)
            ?? throw new InvalidOperationException($"{PowerShellRoutesFile} is empty.");
        using JsonDocument snapshot = JsonDocument.Parse(File.ReadAllText(snapshotPath));
        var commands = new HashSet<string>(
            snapshot.RootElement.GetProperty("commands").EnumerateArray().Select(static command => command.GetProperty("name").GetString()!),
            StringComparer.OrdinalIgnoreCase);
        var routes = new Dictionary<string, PowerShellRoute>(StringComparer.Ordinal);
        foreach (PowerShellRoute route in listed) {
            if (string.IsNullOrWhiteSpace(route.Route) || string.IsNullOrWhiteSpace(route.Cmdlet) || string.IsNullOrWhiteSpace(route.Example)) {
                throw new InvalidOperationException($"{PowerShellRoutesFile} has an entry without a route, cmdlet and example: {route.Route}");
            }
            if (!commands.Contains(route.Cmdlet)) {
                throw new InvalidOperationException($"{PowerShellRoutesFile}: route '{route.Route}' uses '{route.Cmdlet}', which is not in {PowerShellSnapshot}.");
            }
            if (!routes.TryAdd(route.Route, route)) {
                throw new InvalidOperationException($"{PowerShellRoutesFile} lists route '{route.Route}' twice.");
            }
        }
        return routes;
    }

    /// <summary>Presentation groups for the map. A format that isn't listed lands in "other" until someone places it.</summary>
    /// <summary>Families in the order the maps draw them (the home page places them around a circle).</summary>
    private static readonly string[] Families = ["documents", "spreadsheets", "presentations", "fixed", "images", "books", "email", "bibliography", "markup", "web", "other"];

    /// <summary>Within a family, the best-known formats come first.</summary>
    private static readonly string[] Prominence = [
        "DOCX", "DOC", "ODT", "RTF", "Pages", "Google Docs",
        "XLSX", "ODS", "CSV", "Numbers", "Google Sheets",
        "PPTX", "ODP", "Keynote", "Google Slides",
        "PDF", "XPS/OpenXPS", "MHTML", "Visio", "ODG/FODG",
        "PNG", "SVG", "JPEG", "TIFF", "WebP",
        "OneNote", "EPUB", "Book project",
        "Email", "EML", "MSG", "OFT", "TNEF",
        "BibTeX", "BibLaTeX", "CSL JSON", "RIS", "NBIB/MEDLINE", "EndNote XML",
        "Markdown", "AsciiDoc", "LaTeX", "Plain text", "OfficeIMO Markup",
        "HTML", "Confluence", "ADF"
    ];

    private static int Rank(string format) {
        int index = Array.IndexOf(Prominence, format);
        return index < 0 ? int.MaxValue : index;
    }

    private static string Family(string format) => format switch {
        "DOCX" or "DOC" or "ODT" or "RTF" or "Pages" or "Google Docs" => "documents",
        "XLSX" or "ODS" or "CSV" or "Numbers" or "Google Sheets" => "spreadsheets",
        "PPTX" or "ODP" or "Keynote" or "Google Slides" => "presentations",
        "PDF" or "XPS/OpenXPS" or "MHTML" or "Visio" or "ODG/FODG" => "fixed",
        "PNG" or "SVG" or "JPEG" or "TIFF" or "WebP" => "images",
        "OneNote" or "EPUB" or "Book project" => "books",
        "Email" or "EML" or "MSG" or "OFT" or "TNEF" => "email",
        "BibTeX" or "BibLaTeX" or "CSL JSON" or "RIS" or "NBIB/MEDLINE" or "EndNote XML" => "bibliography",
        "Markdown" or "AsciiDoc" or "LaTeX" or "Plain text" or "OfficeIMO Markup" => "markup",
        "HTML" or "Confluence" or "ADF" => "web",
        _ => "other"
    };
}

internal sealed record FormatMapModel(int SchemaVersion, IReadOnlyList<string> Families, IReadOnlyList<FormatMapFormat> Formats, IReadOnlyList<FormatMapRoute> Routes);

internal sealed record FormatMapFormat(string Name, string Family);

internal sealed record FormatMapRoute(
    string Id, string Source, string Target, string Package, string Support, string Fidelity, string Api, IReadOnlyList<string> Surfaces,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? Cmdlet = null,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? Example = null);

/// <summary>One reviewed entry of powershell-routes.json. Evidence points at the PSWriteOffice source that performs the conversion.</summary>
internal sealed record PowerShellRoute(string Route, string Cmdlet, string Example, string? Evidence = null);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, WriteIndented = true)]
[JsonSerializable(typeof(FormatMapModel))]
[JsonSerializable(typeof(PowerShellRoute[]))]
internal sealed partial class FormatMapJsonContext : JsonSerializerContext;
