using System.Diagnostics.CodeAnalysis;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Markdown;

public partial class MarkdownDocument {
    /// <summary>Adds a table from declared properties, fields, or string-keyed dictionary rows in a single pass.</summary>
    /// <remarks>Typed projection rejects tables exceeding 10,000 data rows or 100 columns before adding a block.</remarks>
    public TableBlock AddTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
        IEnumerable<T> rows, OfficeTableOptions? options = null) {
        ValidateTableOptions(options);
        var data = ObjectDataHelpers.ProjectTypedTable(rows, options?.IncludeHeaders ?? true,
            maximumRows: 10000, maximumColumns: 100, options: options);
        var builder = new MarkdownTableBuilder();
        if (options?.IncludeHeaders ?? true) builder.Headers(data.Headers);
        foreach (var row in data.Rows)
            builder.Row(row.Select(value => MarkdownTableBuilder.FormatValue(value)).ToArray());
        return AddTableBlock(builder, options);
    }

    /// <summary>Adds a table from runtime data, including scalar sequences and property/value objects.</summary>
    [RequiresUnreferencedCode("Inspects arbitrary runtime types. Use AddTable<T> with declared rows or explicit selectors for trimming.")]
    [RequiresDynamicCode("Inspects arbitrary runtime dictionaries. Use AddTable<T> with declared rows or explicit selectors for NativeAOT.")]
    public TableBlock AddTable(object? data, OfficeTableOptions? options = null) {
        ValidateTableOptions(options);
        var builder = new MarkdownTableBuilder().FromAny(data, options);
        if (options?.IncludeHeaders == false) builder.Build().Headers.Clear();
        return AddTableBlock(builder, options);
    }

    /// <summary>Adds a table using explicit selectors, with no runtime member discovery.</summary>
    public TableBlock AddTable<T>(IEnumerable<T> rows, params (string Header, Func<T, object?> Selector)[] columns) {
        var table = new MarkdownTableBuilder().FromSequence(rows, columns).Build();
        Add(table);
        return table;
    }

    private TableBlock AddTableBlock(MarkdownTableBuilder builder, OfficeTableOptions? options) {
        if (options is MarkdownTableOptions markdown) {
            if (markdown.AlignDates) builder.AlignDatesCenter();
            if (markdown.AlignNumeric) builder.AlignNumericRight();
        }
        // Explicit column alignment takes precedence over heuristics.
        if (options?.Alignments.Count > 0) builder.Align(options.Alignments.Select(MarkdownTableBuilder.ToMarkdownAlignment).ToArray());
        TableBlock table = builder.Build();
        Add(table);
        return table;
    }

    private static void ValidateTableOptions(OfficeTableOptions? options) {
        ObjectDataHelpers.ValidateTableAlignments(options);
        if (options == null) return;
        foreach (var alignment in options.Alignments) _ = MarkdownTableBuilder.ToMarkdownAlignment(alignment);
    }
}
