using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

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
            Columns? columns = _sectionProperties.GetFirstChild<Columns>();
            if (columns?.EqualWidth?.Value != false) return Array.Empty<WordSectionColumn>();
            var definitions = new List<WordSectionColumn>();
            foreach (Column column in columns.Elements<Column>()) {
                if (!int.TryParse(column.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int width) || width <= 0)
                    throw new InvalidDataException("An explicit section column has an invalid width.");
                int? space = null;
                if (column.Space != null) {
                    if (!int.TryParse(column.Space.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int gap) || gap < 0)
                        throw new InvalidDataException("An explicit section column has an invalid following gap.");
                    space = gap;
                }
                definitions.Add(new WordSectionColumn(width, space));
            }
            return definitions.AsReadOnly();
        }
        set {
            if (value == null) throw new ArgumentNullException(nameof(value));
            WordSectionColumn[] snapshot = value.ToArray();
            if (snapshot.Length > short.MaxValue || snapshot.Any(column => column == null))
                throw new ArgumentException("Column definitions must contain non-null columns within the section count range.", nameof(value));
            Columns? columns = _sectionProperties.GetFirstChild<Columns>();
            if (snapshot.Length == 0 && columns == null) return;
            int? currentCount = ColumnCount;
            if (columns == null) {
                columns = new Columns();
                _sectionProperties.Append(columns);
            }
            columns.RemoveAllChildren<Column>();
            columns.EqualWidth = snapshot.Length == 0;
            columns.ColumnCount = (short)(snapshot.Length == 0 ? currentCount ?? 1 : snapshot.Length);
            foreach (WordSectionColumn definition in snapshot) {
                columns.Append(new Column {
                    Width = definition.WidthTwips.ToString(CultureInfo.InvariantCulture),
                    Space = definition.SpaceAfterTwips.HasValue
                        ? new DocumentFormat.OpenXml.StringValue(definition.SpaceAfterTwips.Value.ToString(CultureInfo.InvariantCulture))
                        : null
                });
            }
        }
    }
}
