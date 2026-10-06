using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write;

internal static partial class LegacyDocWriter {
    private const int MaxLegacySectionColumns = LegacyDocSectionColumns.MaximumCount;

    private static IReadOnlyList<WordSectionColumn>? ReadSupportedColumns(
        Columns columns, out int? columnCount, out int? columnSpacing, out bool hasColumnSeparator) {
        columnCount = ReadColumnCount(columns.ColumnCount);
        columnSpacing = ReadColumnSpacing(columns.Space, columnCount != null ? DefaultColumnSpaceTwips : null);
        hasColumnSeparator = columns.Separator?.Value ?? false;
        if (columns.EqualWidth?.Value != false) return null;
        Column[] explicitColumns = columns.Elements<Column>().ToArray();
        if (explicitColumns.Length < 1 || explicitColumns.Length > MaxLegacySectionColumns ||
            (columnCount.HasValue && columnCount.Value != explicitColumns.Length))
            throw new NotSupportedException("Native DOC saving requires one explicit width per unequal section column.");
        columnCount = explicitColumns.Length;
        var definitions = new List<WordSectionColumn>();
        foreach (Column column in explicitColumns) {
            int width = ReadTwipValue(column.Width, 0, "explicit section column width") ?? 0;
            int gap = ReadColumnSpacing(column.Space, 0) ?? 0;
            if (width < LegacyDocSectionColumns.MinimumWidth || width > LegacyDocSectionColumns.MaximumDimension ||
                gap > LegacyDocSectionColumns.MaximumDimension)
                throw new NotSupportedException("Native DOC explicit columns require widths from 718 through 32767 twips and gaps from zero through 32767 twips.");
            definitions.Add(new WordSectionColumn(width, gap));
        }
        return definitions.AsReadOnly();
    }

    private static int? ReadColumnCount(OpenXmlSimpleType? value) {
        if (value == null) return null;
        if (!int.TryParse(value.InnerText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
            || actual < 1 || actual > MaxLegacySectionColumns)
            throw new NotSupportedException($"Native DOC saving supports section column counts from 1 through {MaxLegacySectionColumns}.");
        return actual;
    }

    private static int? ReadColumnSpacing(OpenXmlSimpleType? value, int? defaultValue) {
        if (value == null) return defaultValue;
        if (!int.TryParse(value.InnerText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
            || actual < 0 || actual > ushort.MaxValue)
            throw new NotSupportedException("Native DOC saving supports section column spacing only within the Word 97-2003 unsigned twip range.");
        return actual;
    }

    private static void AddSectionColumnDefinitions(List<byte> grpprl, IReadOnlyList<WordSectionColumn>? definitions) {
        if (definitions == null || definitions.Count == 0) return;
        AddSingleByteSprm(grpprl, LegacyDocSectionColumns.EvenlySpacedSprm, 0);
        for (int index = 0; index < definitions.Count; index++) {
            AddIndexedColumnSprm(grpprl, LegacyDocSectionColumns.WidthSprm, index, definitions[index].WidthTwips);
            AddIndexedColumnSprm(grpprl, LegacyDocSectionColumns.SpacingSprm, index, definitions[index].SpaceAfterTwips ?? 0);
        }
    }

    private static void AddIndexedColumnSprm(List<byte> grpprl, ushort sprm, int index, int value) {
        grpprl.Add((byte)sprm);
        grpprl.Add((byte)(sprm >> 8));
        grpprl.Add((byte)index);
        grpprl.Add((byte)value);
        grpprl.Add((byte)(value >> 8));
    }
}
