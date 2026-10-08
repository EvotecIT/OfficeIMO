using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>A normalized subpath retaining analytic commands and explicit closure.</summary>
internal sealed class OfficePathContour {
    internal IReadOnlyList<OfficePathCommand> Commands { get; }
    internal bool IsClosed => Commands[Commands.Count - 1].Kind == OfficePathCommandKind.Close;

    private OfficePathContour(List<OfficePathCommand> commands) => Commands = commands.AsReadOnly();

    /// <summary>Splits at moves and closes, retaining empty contours and the current point after a close.</summary>
    internal static IReadOnlyList<OfficePathContour> Split(IReadOnlyList<OfficePathCommand> commands) {
        var contours = new List<OfficePathContour>();
        List<OfficePathCommand>? current = null;
        OfficePoint start = default;
        foreach (OfficePathCommand command in commands) {
            if (command.Kind == OfficePathCommandKind.MoveTo) {
                if (current != null) contours.Add(new OfficePathContour(current));
                current = new List<OfficePathCommand> { command }; start = command.Point;
            } else if (command.Kind == OfficePathCommandKind.Close) {
                if (current == null) continue;
                current.Add(command); contours.Add(new OfficePathContour(current)); current = null;
            } else {
                // Drawing after Z begins at the preceding subpath's starting point.
                current ??= new List<OfficePathCommand> { OfficePathCommand.MoveTo(start) };
                current.Add(command);
            }
        }
        if (current != null) contours.Add(new OfficePathContour(current));
        return contours.AsReadOnly();
    }

    /// <summary>Opens a closed contour while retaining its closing edge, without flattening curves.</summary>
    internal IReadOnlyList<OfficePathCommand> OpenWithClosingEdge() {
        if (!IsClosed) return Commands;
        var open = new List<OfficePathCommand>(Commands.Count);
        for (int i = 0; i < Commands.Count - 1; i++) open.Add(Commands[i]);
        if (open.Count > 1 && open[open.Count - 1].Point != open[0].Point)
            open.Add(OfficePathCommand.LineTo(open[0].Point));
        return open.AsReadOnly();
    }
}
