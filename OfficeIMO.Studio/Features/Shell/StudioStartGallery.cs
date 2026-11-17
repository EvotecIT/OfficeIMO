using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>
/// The start screen's tool gallery: catalog commands grouped like the public Studio feature matrix, so a first look
/// shows what Studio does. Titles, descriptions, availability and execution all come from the command catalog.
/// </summary>
public sealed class StudioStartGallery {
    /// <summary>Gallery groups in matrix order, with the catalog commands each one shows.</summary>
    internal static IReadOnlyList<(string Group, string Title, string[] Commands)> Layout { get; } = [
        ("Read", "Read and find", ["Read", "FocusReading", "PageNightMode"]),
        ("Markup", "Mark up and review", ["Comment", "Tool.Highlight", "CompareSideBySide"]),
        ("Pages", "Organize pages", ["Pages", "Split", "ImportPages", "InsertBlankPage", "PageNumbers", "Watermark"]),
        ("Edit", "Edit content", ["Edit", "Tool.AddImage", "Tool.Link"]),
        ("Forms", "Forms", ["Forms", "CreateSignature", "PlaceDate", "FlattenForms"]),
        ("Scan", "Scan and recognize text", ["Ocr", "DetectPageOrientation"]),
        ("Convert", "Convert", ["Convert", "Export"]),
        ("Security", "Protect and sign", ["Protect", "Redact", "ValidateSignatures", "BatesNumbers"]),
        ("Output", "Print and automate", ["Print", "Jobs"]),
        ("Workbenches", "Workbenches for larger jobs", ["Assemble", "Inspect", "Compare", "Optimize", "Repair", "Sanitize", "PublishBook", "Invoices", "Provenance"])
    ];

    private readonly Dictionary<string, StudioStartTile> _tiles = new(StringComparer.Ordinal);

    internal StudioStartGallery(MainWindowViewModel document, StudioCommandCatalog catalog, IStudioLocalizer localizer) {
        Groups = Layout.Select(group => new StudioStartGroup(
            localizer.GetOrDefault("Home.Gallery." + group.Group, group.Title),
            group.Commands.Select(id => Tile(document, catalog[id])).ToArray())).ToArray();
        // Quick tasks use the same tiles, so a document task on Home asks for a PDF the same way.
        foreach (string id in new[] { "Assemble", "Convert", "Ocr", "Optimize", "Protect", "Tools" })
            Tile(document, catalog[id]);
    }

    public IReadOnlyList<StudioStartGroup> Groups { get; }

    /// <summary>The tile for a catalog command shown on the start screen.</summary>
    public StudioStartTile this[string id] => _tiles[id];

    private StudioStartTile Tile(MainWindowViewModel document, StudioCommandItem command) {
        if (!_tiles.TryGetValue(command.Id, out StudioStartTile? tile))
            _tiles.Add(command.Id, tile = new StudioStartTile(document, command));
        return tile;
    }
}

/// <summary>A titled set of start-screen tiles.</summary>
public sealed record StudioStartGroup(string Title, IReadOnlyList<StudioStartTile> Tiles);

/// <summary>
/// A start-screen entry for a catalog command. A document command with no document open stays usable: it asks for a
/// PDF and then runs on the document that opens.
/// </summary>
public sealed class StudioStartTile : ObservableObject, ICommand {
    private readonly MainWindowViewModel _document;

    internal StudioStartTile(MainWindowViewModel document, StudioCommandItem command) {
        _document = document;
        Command = command;
        command.CanExecuteChanged += (_, _) => {
            OnPropertyChanged(nameof(OpensDocumentFirst));
            OnPropertyChanged(nameof(HelpText));
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        };
    }

    public StudioCommandItem Command { get; }
    public string Id => Command.Id;
    public string Title => Command.Title;
    public string Description => Command.Description;
    /// <summary>True when choosing the tile first asks for a PDF to open.</summary>
    public bool OpensDocumentFirst => Command.RequiresDocument && !_document.HasDocument;
    /// <summary>The tooltip and accessible help: the description, and that a PDF opens first when it does.</summary>
    public string HelpText => OpensDocumentFirst
        ? Description + " " + _document.PaneLocalizer.Get("Home.Gallery.OpensPdfFirst") + "."
        : Description;

    public event EventHandler? CanExecuteChanged;

    public bool CanExecute(object? parameter) => Command.IsAvailable || (OpensDocumentFirst && _document.CanStartDocumentTransition);

    public async void Execute(object? parameter) => await _document.StartToolAsync(Command.Id);
}
