using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.IWork;

/// <summary>One Numbers drawable retained in sheet source order.</summary>
public sealed class IWorkNumbersDrawable {
    internal IWorkNumbersDrawable(IWorkTable table) {
        Kind = IWorkNumbersDrawableKind.Table;
        Table = table;
    }

    internal IWorkNumbersDrawable(string textBox) {
        Kind = IWorkNumbersDrawableKind.TextBox;
        TextBox = textBox;
    }

    /// <summary>Gets the drawable kind.</summary>
    public IWorkNumbersDrawableKind Kind { get; }
    /// <summary>Gets the table when <see cref="Kind"/> is <see cref="IWorkNumbersDrawableKind.Table"/>.</summary>
    public IWorkTable? Table { get; }
    /// <summary>Gets the text when <see cref="Kind"/> is <see cref="IWorkNumbersDrawableKind.TextBox"/>.</summary>
    public string? TextBox { get; }
}

/// <summary>One Numbers sheet and its semantic drawables.</summary>
public sealed class IWorkNumbersSheet {
    internal IWorkNumbersSheet(string name, IReadOnlyList<IWorkTable> tables,
        IReadOnlyList<string> textBoxes, IReadOnlyList<IWorkNumbersDrawable>? drawables = null) {
        Name = name;
        Tables = System.Array.AsReadOnly(tables.ToArray());
        TextBoxes = System.Array.AsReadOnly(textBoxes.ToArray());
        IEnumerable<IWorkNumbersDrawable> ordered = drawables
            ?? textBoxes.Select(text => new IWorkNumbersDrawable(text))
                .Concat(tables.Select(table => new IWorkNumbersDrawable(table)));
        Drawables = System.Array.AsReadOnly(ordered.ToArray());
    }

    /// <summary>Gets the source sheet name.</summary>
    public string Name { get; }
    /// <summary>Gets tables in drawable order.</summary>
    public IReadOnlyList<IWorkTable> Tables { get; }
    /// <summary>Gets text-box content in drawable order.</summary>
    public IReadOnlyList<string> TextBoxes { get; }
    /// <summary>Gets tables and text boxes in their shared source order.</summary>
    public IReadOnlyList<IWorkNumbersDrawable> Drawables { get; }
}
