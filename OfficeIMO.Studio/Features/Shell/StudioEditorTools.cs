using Avalonia.Input;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>A document tool, the mode that offers it, its single-key shortcut inside that mode, and its presentation.</summary>
internal sealed record StudioEditorTool(string Tool, StudioDocumentMode Mode, Key? Key, string NameKey, string IconKey) {
    internal string CommandId => "Tool." + Tool;
}

/// <summary>
/// The annotation and editing tools. The single-key shortcuts, command search entries and their icons all come
/// from this table, so a letter shown next to a tool is the letter that selects it.
/// </summary>
internal static class StudioEditorTools {
    internal static IReadOnlyList<StudioEditorTool> All { get; } = [
        new("Select", StudioDocumentMode.Annotate, Key.V, "DocumentWorkspace.Select", "IconCursor"),
        new("Note", StudioDocumentMode.Annotate, Key.N, "DocumentWorkspace.Note", "IconNote"),
        new("FreeText", StudioDocumentMode.Annotate, Key.T, "DocumentWorkspace.Text", "IconText"),
        new("Highlight", StudioDocumentMode.Annotate, Key.H, "DocumentWorkspace.Highlight", "IconHighlight"),
        new("Underline", StudioDocumentMode.Annotate, Key.U, "DocumentWorkspace.Underline", "IconUnderline"),
        new("StrikeOut", StudioDocumentMode.Annotate, Key.S, "DocumentWorkspace.Strikeout", "IconStrike"),
        new("Squiggly", StudioDocumentMode.Annotate, Key.Q, "Tools.Squiggly", "IconSquiggly"),
        new("Rectangle", StudioDocumentMode.Annotate, Key.R, "DocumentWorkspace.Rectangle", "IconRectangle"),
        new("Ellipse", StudioDocumentMode.Annotate, Key.E, "Tools.Ellipse", "IconEllipse"),
        new("Line", StudioDocumentMode.Annotate, Key.L, "DocumentWorkspace.Line", "IconLine"),
        new("Polygon", StudioDocumentMode.Annotate, Key.P, "Tools.Polygon", "IconPolygon"),
        new("Ink", StudioDocumentMode.Annotate, Key.I, "Tools.Ink", "IconInk"),
        new("Stamp", StudioDocumentMode.Annotate, null, "Tools.Stamp", "IconStamp"),
        new("Select", StudioDocumentMode.Edit, Key.V, "DocumentWorkspace.Select", "IconCursor"),
        new("AddText", StudioDocumentMode.Edit, Key.T, "DocumentWorkspace.AddText", "IconText"),
        new("AddImage", StudioDocumentMode.Edit, null, "DocumentWorkspace.AddImage", "IconImage"),
        new("Link", StudioDocumentMode.Edit, Key.K, "DocumentWorkspace.AddLink", "IconLink")
    ];

    /// <summary>The tool a plain key selects in the current mode, if any.</summary>
    internal static StudioEditorTool? ForKey(StudioDocumentMode mode, Key key) =>
        All.FirstOrDefault(tool => tool.Mode == mode && tool.Key == key);

    /// <summary>Tools reachable from command search; selection exists in every mode and stays on the toolbar.</summary>
    internal static IEnumerable<StudioEditorTool> Searchable => All.Where(tool => tool.Tool != "Select");

    internal static StudioEditorTool? ForCommand(string? commandId) =>
        commandId is { Length: > 5 } id && id.StartsWith("Tool.", StringComparison.Ordinal)
            ? Searchable.FirstOrDefault(tool => tool.CommandId == id)
            : null;
}
