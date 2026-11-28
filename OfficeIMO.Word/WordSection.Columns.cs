using System.Globalization;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordSection {
    /// <summary>
    /// Gets or replaces explicit section column widths and individual gaps.
    /// Setting a non-empty list selects unequal-width columns and synchronizes ColumnCount.
    /// An empty list restores equal-width columns while retaining their count and default gap.
    /// An omitted individual gap is zero for unequal columns; ColumnsSpace applies to equal widths.
    /// DOCX preserves the omission; native DOC writes its effective zero explicitly.
    /// Native DOC saving supports up to 44 columns, widths from 718 through 32767 twips,
    /// and gaps from zero through 32767 twips.
    /// </summary>
    public IReadOnlyList<WordSectionColumn> ColumnDefinitions {
        get {
            OfficeXmlNode columns = SectionNode.Child(SectionNamespace, "cols");
            if (!HasExplicitColumnWidths(columns)) return Array.Empty<WordSectionColumn>();
            var definitions = new List<WordSectionColumn>();
            foreach (OfficeXmlNode column in columns.Children(SectionNamespace, "col")) {
                if (!int.TryParse(column.Attribute(SectionNamespace, "w"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int width) || width <= 0)
                    throw new InvalidDataException("An explicit section column has an invalid width.");
                int? space = null;
                if (column.Attribute(SectionNamespace, "space") is { } spaceToken) {
                    if (!int.TryParse(spaceToken, NumberStyles.Integer, CultureInfo.InvariantCulture, out int gap) || gap < 0)
                        throw new InvalidDataException("An explicit section column has an invalid following gap.");
                    space = gap;
                }
                definitions.Add(new WordSectionColumn(OfficeLength.Points(width / 20D),
                    space.HasValue ? OfficeLength.Points(space.Value / 20D) : (OfficeLength?)null));
            }
            return definitions.AsReadOnly();
        }
        set {
            ArgumentNullException.ThrowIfNull(value);
            WordSectionColumn[] snapshot = value.ToArray();
            if (snapshot.Length > short.MaxValue || snapshot.Any(column => column == null))
                throw new ArgumentException("Column definitions must contain non-null columns within the section count range.", nameof(value));
            OfficeXmlNode columns = SectionNode.Child(SectionNamespace, "cols");
            if (snapshot.Length == 0 && columns.IsNone) return;
            int? currentCount = ColumnCount;
            columns = EditSectionColumns();
            columns.RemoveChildren(SectionNamespace, "col");
            columns.SetAttribute(SectionNamespace, "equalWidth", snapshot.Length == 0 ? "true" : "false");
            columns.SetAttribute(SectionNamespace, "num", (snapshot.Length == 0 ? currentCount ?? 1 : snapshot.Length).ToString(CultureInfo.InvariantCulture));
            foreach (WordSectionColumn definition in snapshot) {
                OfficeXmlNode column = columns.CreateElement(SectionNamespace, "col");
                column.SetAttribute(SectionNamespace, "w", checked((int)definition.Width.ToTwips()).ToString(CultureInfo.InvariantCulture));
                if (definition.SpaceAfter is { } space) {
                    column.SetAttribute(SectionNamespace, "space", checked((int)space.ToTwips()).ToString(CultureInfo.InvariantCulture));
                }
                columns.Append(column);
            }
        }
    }

    private OfficeXmlNode EditSectionColumns() {
        OfficeXmlNode section = EditSectionNode();
        OfficeXmlNode columns = section.Child(SectionNamespace, "cols");
        return columns.Exists ? columns : section.Append(section.CreateElement(SectionNamespace, "cols"));
    }

    private static bool HasExplicitColumnWidths(OfficeXmlNode columns) =>
        columns.Attribute(SectionNamespace, "equalWidth") is "false" or "0" or "off";

    private static int CountDeclaredColumns(OfficeXmlNode columns) {
        int count = 0;
        foreach (OfficeXmlNode _ in columns.Children(SectionNamespace, "col")) count++;
        return count;
    }
}
