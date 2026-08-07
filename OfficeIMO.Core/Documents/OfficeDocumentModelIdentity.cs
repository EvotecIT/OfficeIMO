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
    public static string BuildTableIdentity(OfficeDocumentModelTable table, OfficeDocumentModelPage page, int tableIndex) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        OfficeDocumentModelLocation source = page.Location;
        var fallback = new OfficeDocumentModelLocation {
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
        int? fallbackTableIndex) {
        if (table == null) throw new ArgumentNullException(nameof(table));
        var builder = new StringBuilder();
        Append(builder, table.PayloadHash);
        Append(builder, table.Kind);
        Append(builder, table.Title);
        AppendLocation(builder, table.Location, fallback, fallbackTableIndex);
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
