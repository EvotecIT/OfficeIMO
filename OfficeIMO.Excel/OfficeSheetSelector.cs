using System.Globalization;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Excel;

/// <summary>An immutable worksheet name, wildcard pattern, one-based position, or live worksheet identity.</summary>
public sealed class OfficeSheetSelector {
    /// <summary>Creates a persistable name/pattern or position selector. Exactly one value is required.</summary>
    [JsonConstructor]
    public OfficeSheetSelector(string? name = null, int? position = null) {
        if ((name != null) == position.HasValue) throw new ArgumentException("Specify either a sheet name or a one-based position.");
        if (name != null && string.IsNullOrWhiteSpace(name)) throw new ArgumentException("A sheet name cannot be empty.", nameof(name));
        if (position.HasValue && position.Value < 1) throw new ArgumentOutOfRangeException(nameof(position), "Sheet positions start at 1.");
        Name = name;
        Position = position;
    }

    private OfficeSheetSelector(ExcelSheet sheet) {
        Name = sheet.Name;
        Sheet = sheet;
    }

    /// <summary>Gets the name or wildcard pattern. A numeric string remains a name. Live selectors capture the name for persistence.</summary>
    public string? Name { get; }
    /// <summary>Gets the one-based worksheet position, resolved when the workbook is available.</summary>
    public int? Position { get; }
    /// <summary>Gets the live worksheet identity. JSON persists its captured name without serializing the document.</summary>
    [JsonIgnore]
    public ExcelSheet? Sheet { get; }

    /// <summary>Creates a selector from a string, integral position, existing selector, or worksheet. Hosts unwrap their own argument wrappers first.</summary>
    public static OfficeSheetSelector From(object value) {
        ArgumentNullException.ThrowIfNull(value);
        if (value is OfficeSheetSelector selector) return selector;
        if (value is string name) return new OfficeSheetSelector(name);
        if (value is ExcelSheet sheet) return new OfficeSheetSelector(sheet);
        if (value is byte or sbyte or short or ushort or int or uint or long or ulong) {
            try { return new OfficeSheetSelector(position: checked(Convert.ToInt32(value, CultureInfo.InvariantCulture))); }
            catch (OverflowException exception) { throw new ArgumentException("Sheet positions must fit an Int32.", nameof(value), exception); }
        }
        throw new ArgumentException("Expected a sheet name, a one-based integer or an ExcelSheet.", nameof(value));
    }

    /// <summary>Resolves a literal name, position or live identity against a worksheet collection. Use ResolveMany for wildcard expansion.</summary>
    public T Resolve<T>(IReadOnlyList<T> sheets, Func<T, string> getName) where T : class {
        ArgumentNullException.ThrowIfNull(sheets);
        ArgumentNullException.ThrowIfNull(getName);
        if (Position.HasValue) {
            if (Position.Value <= sheets.Count) return sheets[Position.Value - 1];
        } else {
            foreach (T item in sheets) {
                bool matches = Sheet == null
                    ? string.Equals(getName(item), Name, StringComparison.OrdinalIgnoreCase)
                    : ReferenceEquals(item, Sheet) || item is ExcelSheet candidate
                        && ReferenceEquals(candidate.Document, Sheet.Document)
                        && ReferenceEquals(candidate.SheetElement, Sheet.SheetElement);
                if (matches) return item;
            }
        }
        throw new ArgumentException("The selected sheet does not belong to this document.");
    }

    /// <summary>
    /// Resolves selectors once against the loaded workbook, deduplicating in workbook order.
    /// Names support case-insensitive *, ?, bracket sets/ranges and backtick escapes. Unmatched or empty explicit selections are errors.
    /// Null selects the supplied default worksheet or all worksheets.
    /// </summary>
    public static IReadOnlyList<ExcelSheet> ResolveMany(ExcelDocument document,
        IReadOnlyList<OfficeSheetSelector>? selectors, ExcelSheet? defaultSheet = null) {
        ArgumentNullException.ThrowIfNull(document);
        IReadOnlyList<ExcelSheet> sheets = document.Sheets;
        if (selectors == null && defaultSheet != null)
            return new[] { From(defaultSheet).Resolve(sheets, sheet => sheet.Name) };
        return ResolveSelection(sheets, sheet => sheet.Name, selectors, allowLiveSheet: true);
    }

    /// <summary>
    /// Resolves the same name/pattern and position selectors against an existing reader's worksheet names,
    /// without loading an editable workbook. Live worksheet identities require ResolveMany instead.
    /// Null selects all names; explicit empty or unmatched selections are errors. Results follow workbook order.
    /// </summary>
    public static IReadOnlyList<string> ResolveNames(IReadOnlyList<string> sheetNames,
        IReadOnlyList<OfficeSheetSelector>? selectors) {
        ArgumentNullException.ThrowIfNull(sheetNames);
        return ResolveSelection(sheetNames, name => name, selectors, allowLiveSheet: false);
    }

    private static IReadOnlyList<T> ResolveSelection<T>(IReadOnlyList<T> sheets, Func<T, string> getName,
        IReadOnlyList<OfficeSheetSelector>? selectors, bool allowLiveSheet) where T : class {
        if (selectors == null) return sheets.ToArray();
        if (selectors.Count == 0) throw new ArgumentException("At least one worksheet selector is required.", nameof(selectors));
        var selected = new HashSet<T>();
        foreach (OfficeSheetSelector selector in selectors) {
            if (selector == null) throw new ArgumentException("Worksheet selectors cannot contain null.", nameof(selectors));
            if (!allowLiveSheet && selector.Sheet != null)
                throw new ArgumentException("Live worksheet selectors require the owning loaded workbook.", nameof(selectors));
            if (selector.Sheet != null || selector.Position.HasValue || !ContainsNameWildcard(selector.Name!)) {
                selected.Add(selector.Resolve(sheets, getName));
                continue;
            }
            var matcher = CreateNameMatcher(selector.Name!);
            bool found = false;
            foreach (T sheet in sheets) {
                if (!matcher.IsMatch(getName(sheet))) continue;
                selected.Add(sheet);
                found = true;
            }
            if (!found) throw new ArgumentException($"No worksheet matches '{selector.Name}'.", nameof(selectors));
        }
        return sheets.Where(selected.Contains).ToArray();
    }

    private static bool ContainsNameWildcard(string name) {
        for (int index = 0; index < name.Length; index++) {
            if (name[index] == '`' && index + 1 < name.Length) index++;
            else if (name[index] is '*' or '?' or '[') return true;
        }
        return false;
    }

    private static Regex CreateNameMatcher(string pattern) {
        var expression = new StringBuilder("\\A");
        for (int index = 0; index < pattern.Length; index++) {
            char current = pattern[index];
            if (current == '`') {
                expression.Append(Regex.Escape((index + 1 < pattern.Length ? pattern[++index] : current).ToString()));
            } else if (current == '*') expression.Append(".*");
            else if (current == '?') expression.Append('.');
            else if (current == '[') {
                expression.Append('[');
                bool closed = false;
                while (++index < pattern.Length) {
                    current = pattern[index];
                    if (current == ']') { closed = true; expression.Append(']'); break; }
                    bool escaped = current == '`' && index + 1 < pattern.Length;
                    if (escaped) current = pattern[++index];
                    if (current is '\\' or ']' or '^' or '[' || escaped && current == '-') expression.Append('\\');
                    expression.Append(current);
                }
                if (!closed) throw new ArgumentException("A worksheet wildcard bracket set is not closed.", nameof(pattern));
            } else expression.Append(Regex.Escape(current.ToString()));
        }
        expression.Append("\\z");
        return new Regex(expression.ToString(), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline,
            TimeSpan.FromSeconds(1));
    }
}
