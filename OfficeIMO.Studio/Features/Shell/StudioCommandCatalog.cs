using System.ComponentModel;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Shared discovery, labels, and availability for Studio commands; document behavior stays in the view model and engines.</summary>
public sealed class StudioCommandCatalog : ObservableObject, IDisposable {
    private readonly MainWindowViewModel _document;
    private readonly Dictionary<string, StudioCommandItem> _byId = new(StringComparer.Ordinal);
    private string _toolQuery = string.Empty;

    internal StudioCommandCatalog(MainWindowViewModel document, IStudioLocalizer localizer) {
        _document = document;
        string Text(string key, string fallback) => localizer.GetOrDefault("Commands." + key, fallback);
        string? Idle() => document.CanStartDocumentTransition ? null : Text("Busy", "Wait for the current document operation or cancel it.");
        string? Loaded() => Idle() ?? (!document.HasDocument ? Text("OpenFirst", "Open a PDF to use this command.") : null);
        string? Allowed(bool value) => Loaded() ?? (!value ? Text("Unsupported", "This document does not allow this operation. Inspect its protection and supported content.") : null);
        string modifier = OperatingSystem.IsMacOS() ? "⌘" : "Ctrl+";
        void Add(string id, string title, string description, string category, ICommand operation,
            Func<string?>? guard = null, bool tool = false, string shortcut = "", bool workspace = false) {
            _byId.Add(id, new StudioCommandItem(id, Text(id + ".Title", title), Text(id + ".Description", description),
                Text("Category." + category, category), shortcut, tool, operation, guard ?? Idle,
                workspace ? () => { document.IsFocusReading = false; document.WorkspaceMode = StudioWorkspaceMode.PdfWorkspace; } : null));
        }

        Add("Open", "Open PDF", "Open a PDF in a document tab.", "File", document.OpenCommand, shortcut: modifier + "O");
        Add("Save", "Save", "Save changes to the current document.", "File", document.SaveCommand,
            () => Loaded() ?? (!document.IsDirty ? Text("NoChanges", "There are no unsaved changes.") : null), shortcut: modifier + "S");
        Add("SaveAs", "Save a copy", "Choose a destination for this document.", "File", document.SaveAsCommand, Loaded, shortcut: modifier + "Shift+S");
        Add("Undo", "Undo", "Undo the last document edit.", "Edit", document.UndoCommand,
            () => Loaded() ?? (!document.CanUndo ? Text("NoUndo", "There is no edit to undo.") : null), shortcut: modifier + "Z");
        Add("Redo", "Redo", "Restore the last undone edit.", "Edit", document.RedoCommand,
            () => Loaded() ?? (!document.CanRedo ? Text("NoRedo", "There is no edit to redo.") : null), shortcut: modifier + "Shift+Z");
        Add("Read", "Read document", "Return to the document reading workspace.", "Read", document.ShowViewModeCommand, Loaded, workspace: true);
        Add("FocusReading", "Focus reading", "Hide document tools and panes, or restore the workspace.", "Read", document.ToggleFocusReadingCommand, Loaded, shortcut: "F9");
        Add("Comment", "Comment and review", "Add annotations and review existing comments.", "Review", document.ShowAnnotateModeCommand, Loaded, true, workspace: true);
        Add("Edit", "Edit PDF content", "Select supported existing text and images to edit.", "Edit", document.ShowEditModeCommand, () => Allowed(document.CanEditPageContent), true, workspace: true);
        Add("Pages", "Organize pages", "Rotate, crop, duplicate, reorder, import, extract, and split pages.", "Organize", document.ShowPagesModeCommand, Loaded, true, workspace: true);
        Add("Forms", "Fill and edit forms", "Fill fields or author supported AcroForm controls.", "Edit", document.ShowFormsModeCommand, Loaded, true, workspace: true);
        Add("Protect", "Protect and sign", "Inspect protection, sign, or protect a document copy.", "Security", document.ShowProtectModeCommand, Loaded, true, workspace: true);
        Add("Redact", "Redact content", "Mark content for reviewed permanent removal.", "Security", document.BeginRedactionCommand, () => Allowed(document.CanRedact), true, workspace: true);
        Add("Convert", "Convert files", "Convert supported Office, web, markup, PDF, and image files.", "Convert", document.ShowConversionWorkbenchCommand, tool: true);
        Add("Ocr", "Make PDF searchable", "Recognize scanned pages and add a searchable text layer.", "Convert", document.ShowOcrCommand, tool: true);
        Add("Assemble", "Assemble PDF", "Combine PDFs, images, Office files, folders, and ZIP archives.", "Organize", document.ShowAssemblyCommand, tool: true);
        Add("Export", "Export page images", "Export selected PDF pages to image files.", "Output", document.ShowPageExportCommand, tool: true);
        Add("Print", "Preview print sheets", "Inspect planned print sheets and page placement.", "Output", document.ShowPrintPreviewCommand, tool: true, shortcut: modifier + "P");
        Add("Inspect", "Inspect document", "Review structure, security, and document findings.", "Review", document.ShowInspectCommand, tool: true);
        Add("Compare", "Compare PDFs", "Compare two PDFs and inspect the resulting report.", "Review", document.ShowCompareCommand, tool: true);
        Add("Optimize", "Optimize PDF", "Plan and run supported lossless optimization.", "Output", document.ShowOptimizeCommand, tool: true);
        Add("Repair", "Plan PDF repair", "Inspect repair findings before choosing an output.", "Review", document.ShowRepairPlanCommand, tool: true);
        Add("Sanitize", "Sanitize PDF", "Remove forbidden actions and payloads, then verify the output.", "Security", document.ShowSanitizeCommand, tool: true);
        Add("FitPage", "Fit page", "Show the complete current page.", "Read", document.FitPageCommand, Loaded, shortcut: modifier + "0");
        Add("ActualSize", "Actual size", "Display the document at its natural scale.", "Read", document.ActualSizeCommand, Loaded, shortcut: modifier + "1");
        Add("FitWidth", "Fit width", "Fit the current page to the reading area.", "Read", document.FitWidthCommand, Loaded, shortcut: modifier + "2");
        Add("ZoomIn", "Zoom in", "Increase document magnification.", "Read", document.ZoomInCommand, Loaded, shortcut: modifier + "+");
        Add("ZoomOut", "Zoom out", "Decrease document magnification.", "Read", document.ZoomOutCommand, Loaded, shortcut: modifier + "−");
        Add("Home", "Home", "Open recent documents and starting tasks.", "Navigate", document.ShowHomeCommand);
        Add("Tools", "Tools", "Search the document tool catalog.", "Navigate", document.ShowToolsCommand);
        Add("Settings", "Settings", "Change appearance and interface preferences.", "Navigate", document.ShowSettingsCommand);
        Items = _byId.Values.ToArray();
        document.PropertyChanged += OnDocumentChanged;
    }

    public IReadOnlyList<StudioCommandItem> Items { get; }
    public StudioCommandItem this[string id] => _byId[id];
    public string ToolQuery {
        get => _toolQuery;
        set { if (SetProperty(ref _toolQuery, value ?? string.Empty)) { OnPropertyChanged(nameof(FilteredTools)); OnPropertyChanged(nameof(HasTools)); } }
    }
    public IReadOnlyList<StudioCommandItem> FilteredTools => Search(ToolQuery, toolsOnly: true);
    public bool HasTools => FilteredTools.Count > 0;

    public IReadOnlyList<StudioCommandItem> Search(string? query, bool toolsOnly = false) {
        string[] terms = (query ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return Items.Where(item => (!toolsOnly || item.IsTool) && terms.All(term =>
            (item.Id + " " + item.Title + " " + item.Description + " " + item.Category + " " + item.Shortcut)
                .Contains(term, StringComparison.CurrentCultureIgnoreCase))).ToArray();
    }

    private void OnDocumentChanged(object? sender, PropertyChangedEventArgs e) {
        foreach (StudioCommandItem item in Items) item.Refresh();
    }

    public void Dispose() => _document.PropertyChanged -= OnDocumentChanged;
}
