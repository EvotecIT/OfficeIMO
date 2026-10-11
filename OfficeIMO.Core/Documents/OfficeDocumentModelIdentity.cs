using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;

namespace OfficeIMO;

/// <summary>Builds deterministic identities for de-duplicating neutral model elements across aggregate and page collections.</summary>
public static class OfficeDocumentModelIdentity {
    /// <summary>Builds a block identity.</summary>
    public static string BuildBlockIdentity(OfficeDocumentModelBlock block) {
        if (block == null) throw new ArgumentNullException(nameof(block));
        if (!string.IsNullOrWhiteSpace(block.Id)) return "id:" + block.Id;
        OfficeDocumentModelLocation? location = block.Location;
        if (!string.IsNullOrWhiteSpace(location?.BlockAnchor)) return "anchor:" + location!.BlockAnchor!;
        return BuildLocatedIdentity(block, location, block.Kind, block.Text);
    }

    /// <summary>Builds a table identity.</summary>
    public static string BuildTableIdentity(OfficeDocumentModelTable table) =>
        BuildTableIdentity(table, null, null);

    /// <summary>Builds a page-scoped table identity.</summary>
    public static string BuildTableIdentity(OfficeDocumentModelTable table, OfficeDocumentModelPage page, int tableIndex) =>
        BuildPageTableIdentity(table, page, tableIndex);

    // Aggregate/page reconciliation counts equal source occurrences separately. Its
    // key must omit an inferred collection index that the aggregate does not have.
    internal static string BuildTableCollectionIdentity(OfficeDocumentModelTable table, OfficeDocumentModelPage page) {
        var scope = new OfficeDocumentModelLocation {
            Path = page.Location.Path, Sheet = page.Location.Sheet, Slide = page.Location.Slide,
            Page = page.Location.Page ?? page.Number
        };
        return WithTableSpanCoordinates(BuildTableIdentity(table, scope, null), table.Location);
    }

    internal static string BuildTableOccurrenceIdentity(OfficeDocumentModelTable table) =>
        WithTableSpanCoordinates(BuildTableIdentity(table), table.Location);

    // Public identity strings retain their established layout. Projection occurrence
    // matching also distinguishes text spans omitted by those legacy identities.
    private static string WithTableSpanCoordinates(string identity, OfficeDocumentModelLocation? location) {
        if (location?.EndLine == null && location?.NormalizedStartLine == null && location?.NormalizedEndLine == null)
            return identity;
        var builder = new StringBuilder(identity);
        AppendCoordinate(builder, "end-line", location?.EndLine);
        AppendCoordinate(builder, "normalized-start-line", location?.NormalizedStartLine);
        AppendCoordinate(builder, "normalized-end-line", location?.NormalizedEndLine);
        return builder.ToString();
    }

    private static void AppendCoordinate(StringBuilder builder, string name, int? value) {
        if (!value.HasValue) return;
        Append(builder, name);
        Append(builder, value.Value.ToString(CultureInfo.InvariantCulture));
    }

    internal static string BuildTableContentIdentity(OfficeDocumentModelTable table) =>
        BuildTableIdentity(table, null, null, includeLocation: false);

    internal static bool TableLocationMatches(OfficeDocumentModelLocation? aggregate,
        OfficeDocumentModelLocation? table, OfficeDocumentModelPage page) =>
        SameWhenKnown(aggregate?.Path, table?.Path ?? page.Location.Path)
        && SameWhenKnown(aggregate?.Sheet, table?.Sheet ?? page.Location.Sheet)
        && SameWhenKnown(aggregate?.Slide, table?.Slide ?? page.Location.Slide)
        && SameWhenKnown(aggregate?.Page, table?.Page ?? page.Location.Page ?? page.Number)
        && SameWhenKnown(aggregate?.LogicalOrder, table?.LogicalOrder)
        && SameWhenKnown(aggregate?.BlockIndex, table?.BlockIndex)
        && SameWhenKnown(aggregate?.SourceBlockIndex, table?.SourceBlockIndex)
        && SameWhenKnown(aggregate?.TableIndex, table?.TableIndex)
        && SameWhenKnown(aggregate?.StartLine, table?.StartLine)
        && SameWhenKnown(aggregate?.EndLine, table?.EndLine)
        && SameWhenKnown(aggregate?.NormalizedStartLine, table?.NormalizedStartLine)
        && SameWhenKnown(aggregate?.NormalizedEndLine, table?.NormalizedEndLine)
        && SameWhenKnown(aggregate?.HeadingPath, table?.HeadingPath)
        && SameWhenKnown(aggregate?.HeadingSlug, table?.HeadingSlug)
        && SameWhenKnown(aggregate?.SourceBlockKind, table?.SourceBlockKind)
        && SameWhenKnown(aggregate?.BlockAnchor, table?.BlockAnchor)
        && SameWhenKnown(aggregate?.A1Range, table?.A1Range);

    private static bool SameWhenKnown<T>(T? left, T? right) where T : struct =>
        !left.HasValue || !right.HasValue || EqualityComparer<T>.Default.Equals(left.Value, right.Value);

    private static bool SameWhenKnown(string? left, string? right) =>
        string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right) || string.Equals(left, right, StringComparison.Ordinal);

    private static string BuildPageTableIdentity(OfficeDocumentModelTable table, OfficeDocumentModelPage page, int? tableIndex) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        OfficeDocumentModelLocation source = page.Location;
        var fallback = new OfficeDocumentModelLocation {
            LogicalOrder = source.LogicalOrder,
            Path = source.Path,
            BlockIndex = source.BlockIndex,
            SourceBlockIndex = source.SourceBlockIndex,
            StartLine = source.StartLine,
            EndLine = source.EndLine,
            NormalizedStartLine = source.NormalizedStartLine,
            NormalizedEndLine = source.NormalizedEndLine,
            HeadingPath = source.HeadingPath,
            HeadingSlug = source.HeadingSlug,
            SourceBlockKind = source.SourceBlockKind,
            BlockAnchor = source.BlockAnchor,
            Sheet = source.Sheet,
            A1Range = source.A1Range,
            Slide = source.Slide,
            Page = source.Page ?? page.Number,
            TableIndex = source.TableIndex
        };
        return BuildTableIdentity(table, fallback, tableIndex);
    }

    /// <summary>Builds an asset identity.</summary>
    public static string BuildAssetIdentity(OfficeDocumentModelAsset asset) {
        if (asset == null) throw new ArgumentNullException(nameof(asset));
        if (!string.IsNullOrWhiteSpace(asset.Id)) return "id:" + asset.Id;
        if (!string.IsNullOrWhiteSpace(asset.SourceObjectId)) return "source:" + asset.SourceObjectId;
        if (!string.IsNullOrWhiteSpace(asset.PayloadHash)) return "hash:" + asset.PayloadHash;
        OfficeDocumentModelLocation? location = asset.Location;
        if (!string.IsNullOrWhiteSpace(location?.BlockAnchor)) return "anchor:" + location!.BlockAnchor!;
        return BuildLocatedIdentity(asset, location, asset.FileName, asset.MediaType, asset.Kind);
    }

    /// <summary>Builds a link identity.</summary>
    public static string BuildLinkIdentity(OfficeDocumentModelLink link) {
        if (link == null) throw new ArgumentNullException(nameof(link));
        if (!string.IsNullOrWhiteSpace(link.Id)) return "id:" + link.Id;
        OfficeDocumentModelLocation? location = link.Location;
        if (!string.IsNullOrWhiteSpace(location?.BlockAnchor)) return "anchor:" + location!.BlockAnchor!;
        return BuildLocatedIdentity(link, location, link.Uri, link.DestinationName, link.RemoteFile, link.Text);
    }

    /// <summary>Builds a form-field identity.</summary>
    public static string BuildFormIdentity(OfficeDocumentModelFormField form) {
        if (form == null) throw new ArgumentNullException(nameof(form));
        if (!string.IsNullOrWhiteSpace(form.Id)) return "id:" + form.Id;
        OfficeDocumentModelLocation? location = form.Location;
        if (!string.IsNullOrWhiteSpace(location?.BlockAnchor)) return "anchor:" + location!.BlockAnchor!;
        return BuildLocatedIdentity(form, location, form.Name, form.Kind);
    }

    private static string BuildTableIdentity(
        OfficeDocumentModelTable table,
        OfficeDocumentModelLocation? fallback,
        int? fallbackTableIndex,
        bool includeLocation = true) {
        if (table == null) throw new ArgumentNullException(nameof(table));
        var builder = new StringBuilder();
        Append(builder, table.PayloadHash);
        Append(builder, table.Kind);
        Append(builder, table.Title);
        if (includeLocation) AppendLocation(builder, table.Location, fallback, fallbackTableIndex);
        Append(builder, table.Columns);
        foreach (IReadOnlyList<string> row in table.Rows ?? Array.Empty<IReadOnlyList<string>>()) Append(builder, row);
        Append(builder, table.TotalRowCount.ToString(CultureInfo.InvariantCulture));
        return builder.ToString();
    }

    private static string BuildLocatedIdentity<T>(
        T instance,
        OfficeDocumentModelLocation? location,
        params string?[] values) where T : class {
        bool hasLocation = location != null && (
            location.LogicalOrder.HasValue ||
            !string.IsNullOrWhiteSpace(location.Path) ||
            !string.IsNullOrWhiteSpace(location.Sheet) ||
            location.Page.HasValue ||
            location.Slide.HasValue ||
            location.BlockIndex.HasValue ||
            location.SourceBlockIndex.HasValue ||
            location.StartLine.HasValue ||
            location.TableIndex.HasValue);
        if (!hasLocation) return "reference:" + RuntimeHelpers.GetHashCode(instance).ToString(CultureInfo.InvariantCulture);

        var builder = new StringBuilder();
        AppendLocation(builder, location, null, null);
        foreach (string? value in values) Append(builder, value);
        return builder.ToString();
    }

    private static void AppendLocation(
        StringBuilder builder,
        OfficeDocumentModelLocation? location,
        OfficeDocumentModelLocation? fallback,
        int? fallbackTableIndex) {
        Append(builder, Prefer(location?.Path, fallback?.Path));
        Append(builder, Prefer(location?.Sheet, fallback?.Sheet));
        Append(builder, Prefer(location?.A1Range, fallback?.A1Range));
        Append(builder, Prefer(location?.HeadingPath, fallback?.HeadingPath));
        Append(builder, Prefer(location?.HeadingSlug, fallback?.HeadingSlug));
        Append(builder, Prefer(location?.SourceBlockKind, fallback?.SourceBlockKind));
        Append(builder, Prefer(location?.BlockAnchor, fallback?.BlockAnchor));
        Append(builder, (location?.Page ?? fallback?.Page)?.ToString(CultureInfo.InvariantCulture));
        Append(builder, (location?.Slide ?? fallback?.Slide)?.ToString(CultureInfo.InvariantCulture));
        Append(builder, (location?.BlockIndex ?? fallback?.BlockIndex)?.ToString(CultureInfo.InvariantCulture));
        Append(builder, (location?.SourceBlockIndex ?? fallback?.SourceBlockIndex)?.ToString(CultureInfo.InvariantCulture));
        Append(builder, (location?.StartLine ?? fallback?.StartLine)?.ToString(CultureInfo.InvariantCulture));
        Append(builder, (location?.TableIndex ?? fallbackTableIndex ?? fallback?.TableIndex)?.ToString(CultureInfo.InvariantCulture));
        long? logicalOrder = location?.LogicalOrder ?? fallback?.LogicalOrder;
        if (logicalOrder.HasValue) {
            Append(builder, "logical-order");
            Append(builder, logicalOrder.Value.ToString(CultureInfo.InvariantCulture));
        }
    }

    private static string? Prefer(string? value, string? fallback) =>
        string.IsNullOrWhiteSpace(value) ? fallback : value;

    private static void Append(StringBuilder builder, IReadOnlyList<string>? values) {
        if (values == null) {
            Append(builder, (string?)null);
            return;
        }
        Append(builder, values.Count.ToString(CultureInfo.InvariantCulture));
        foreach (string value in values) Append(builder, value);
    }

    private static void Append(StringBuilder builder, string? value) {
        if (value == null) {
            builder.Append("-1:");
            return;
        }
        builder.Append(value.Length.ToString(CultureInfo.InvariantCulture));
        builder.Append(':');
        builder.Append(value);
    }
}
