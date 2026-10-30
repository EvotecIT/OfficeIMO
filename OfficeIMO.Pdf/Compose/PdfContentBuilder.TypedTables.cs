using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeIMO.Core.Internal;
using OfficeIMO.Data;

namespace OfficeIMO.Pdf;

public sealed partial class PdfContentBuilder {
    /// <summary>Adds a table using the shared typed projection and default PDF styling.</summary>
    public PdfContentBuilder Table<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
        IEnumerable<T> rows) => Table(rows, null);

    /// <summary>Adds a table from typed properties, tuples, or string-keyed dictionaries using the shared bounded projection engine.</summary>
    /// <param name="rows">Source rows, enumerated once while this call snapshots the table.</param>
    /// <param name="options">Common column selection, ordering, header, formatting and alignment options.</param>
    /// <remarks>A table name becomes the PDF caption. Logical start/end align left/right; justified columns use PDF's left alignment.</remarks>
    public PdfContentBuilder Table<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
        IEnumerable<T> rows, OfficeTableOptions? options) {
        Guard.NotNull(rows, nameof(rows));
        ObjectDataHelpers.ValidateTableAlignments(options);
        var pdfOptions = options as PdfTableOptions;
        PdfAlign align = pdfOptions?.Align ?? PdfAlign.Left;
        Guard.LeftCenterRightAlign(align, nameof(options), "Table");
        var projectionOptions = pdfOptions?.Projection?.CreateProjectionSnapshot() ?? new ObjectFlattenerOptions();
        ObjectDataHelpers.ConfigureTableProjection(projectionOptions, options);
        bool headers = options?.IncludeHeaders ?? true;
        var projection = new ObjectFlattener().FlattenRows(rows, projectionOptions, "PDF Table", headers ? 1 : 0);
        var paths = projection.Columns;
        if (paths.Count == 0) throw new ArgumentException("No columns could be resolved from the supplied data.", nameof(rows));
        var output = new List<string[]>();
        if (headers) output.Add(paths.Select(path => GetTypedTableHeader(path, projectionOptions)).ToArray());
        int dataRows = 0;
        foreach (var row in projection.Rows) {
            string? collectionPath = projectionOptions.CollectionMode == CollectionMode.ExpandRows
                ? paths.FirstOrDefault(path => row.TryGetValue(path, out var value) && value is IEnumerable && value is not string)
                : null;
            bool expanded = false;
            if (collectionPath != null && row[collectionPath] is IEnumerable collection) {
                int items = 0;
                foreach (object? item in collection) {
                    if (items++ >= projectionOptions.MaxCollectionItems)
                        throw new InvalidDataException("PDF Table nested collection exceeds the configured item limit.");
                    AddRow(row, collectionPath, item);
                    expanded = true;
                }
            }
            if (!expanded) AddRow(row, null, null);
        }
        if (output.Count == 0) throw new ArgumentException("A headerless table requires at least one row.", nameof(rows));
        var effectiveStyle = pdfOptions?.Style?.Clone() ?? _doc.CompositionOptions.DefaultTableStyle ?? new PdfTableStyle();
        effectiveStyle.HeaderRowCount = headers ? 1 : 0;
        effectiveStyle.RepeatHeaderRowCount = headers ? 1 : 0;
        if (options?.Name != null) effectiveStyle.Caption = options.Name;
        if (options != null && options.Alignments.Count > 0) {
            var alignments = new List<PdfColumnAlign>();
            for (int column = 0; column < paths.Count; column++) {
                OfficeAlignment? selected = column < options.Alignments.Count ? options.Alignments[column] : null;
                alignments.Add(selected.HasValue ? selected.Value switch {
                    OfficeAlignment.Center => PdfColumnAlign.Center,
                    OfficeAlignment.Right or OfficeAlignment.End => PdfColumnAlign.Right,
                    _ => PdfColumnAlign.Left
                } : effectiveStyle.Alignments != null && column < effectiveStyle.Alignments.Count
                    ? effectiveStyle.Alignments[column] : PdfColumnAlign.Left);
            }
            effectiveStyle.Alignments = alignments;
        }
        _doc.Table(output, align, effectiveStyle);
        return this;

        void AddRow(Dictionary<string, object?> row, string? expandedPath, object? expandedValue) {
            if (dataRows >= projectionOptions.MaxRows || ((long)dataRows + 1 + (headers ? 1 : 0)) * paths.Count > projectionOptions.MaxCells)
                throw new InvalidDataException("PDF Table expansion exceeds the configured row or cell limit.");
            var cells = new string[paths.Count];
            for (int column = 0; column < paths.Count; column++) {
                string path = paths[column];
                object? value = path == expandedPath ? expandedValue
                    : row.TryGetValue(path, out var projectedValue) ? projectedValue
                    : projectionOptions.DefaultValues.TryGetValue(path, out var defaultValue) ? defaultValue : null;
                cells[column] = Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty;
            }
            output.Add(cells);
            dataRows++;
        }
    }

    private static string GetTypedTableHeader(string path, ObjectFlattenerOptions options) {
        if (options.TableOptions is { } tableOptions && (tableOptions.HeaderRenames.ContainsKey(path) || tableOptions.HeaderTransform != null))
            return ObjectDataHelpers.GetTableHeader(path, tableOptions);
        path = options.GetHeaderPath(path);
        return options.HeaderCase switch {
            HeaderCase.Pascal => string.Concat(path.Split('.').Select(part => char.ToUpperInvariant(part[0]) + part.Substring(1))),
            HeaderCase.Title => string.Join(" ", path.Split('.').Select(part => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(part.ToLowerInvariant()))),
            _ => path
        };
    }
}
