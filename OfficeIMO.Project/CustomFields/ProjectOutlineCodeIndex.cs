using System.Globalization;

namespace OfficeIMO.Project;

/// <summary>Validated, bounded hierarchy for a single outline lookup table.</summary>
internal sealed class ProjectOutlineCodeIndex {
    private readonly ProjectOutlineCodeDefinition _definition;
    private readonly Dictionary<int, ProjectOutlineCodeLookupValue> _values = new();
    private readonly HashSet<int> _parents = new();
    private readonly ProjectOutlineCodeMask[] _masks;
    internal ProjectOutlineCodeIndex(ProjectOutlineCodeDefinition definition, CancellationToken token = default) {
        _definition = definition;
        token.ThrowIfCancellationRequested();
        if (definition.Masks.Count > 64) throw new InvalidDataException("Outline-code masks exceed 64 levels.");
        _masks = definition.Masks.OrderBy(m => m.Level).ToArray();
        if (_masks.Length > 64) throw new InvalidDataException("Outline-code masks exceed 64 levels.");
        for (int i = 0; i < _masks.Length; i++) {
            var mask = _masks[i];
            if (mask.Level != i + 1 || mask.Type < 0 || mask.Type > 3 || !mask.Type.HasValue || !mask.Length.HasValue || mask.Length < 0 || mask.Length > 4096)
                throw new InvalidDataException("Outline-code masks require consecutive levels, a supported character type and explicit nonnegative bounded lengths.");
        }
        foreach (var value in definition.Values) {
            token.ThrowIfCancellationRequested();
            if (!value.ValueId.HasValue || value.ValueId <= 0 || _values.ContainsKey(value.ValueId.Value))
                throw new InvalidDataException("Outline-code lookup values require unique positive identities.");
            _values.Add(value.ValueId.Value, value);
            if (value.ParentValueId > 0) _parents.Add(value.ParentValueId.Value);
            if (value.ParentValueId < 0) throw new InvalidDataException("Outline-code parent identities cannot be negative.");
        }
        foreach (int parent in _parents)
            if (!_values.ContainsKey(parent)) throw new InvalidDataException("An outline-code parent identity is missing.");
        var guids = new HashSet<Guid>();
        foreach (var value in definition.Values) {
            token.ThrowIfCancellationRequested();
            if (value.Guid != null && (!Guid.TryParse(value.Guid, out var guid) || !guids.Add(guid)))
                throw new InvalidDataException("Outline-code value GUIDs must be valid and unique.");
            _ = Chain(value);
        }
    }
    internal ProjectOutlineCodeLookupValue Resolve(ProjectCustomFieldValue reference) {
        if (!int.TryParse(reference.ValueId, NumberStyles.None, CultureInfo.InvariantCulture, out int id) || !_values.TryGetValue(id, out var value))
            throw new InvalidDataException("An outline-code selection has no matching lookup value.");
        if (reference.ValueGuid != null && !SameGuid(reference.ValueGuid, value.Guid))
            throw new InvalidDataException("An outline-code selection's ID and GUID disagree.");
        return value;
    }
    internal string Text(ProjectOutlineCodeLookupValue value, bool selection) {
        var chain = Chain(value);
        if (selection && _definition.LeafOnly == true && _parents.Contains(value.ValueId!.Value))
            throw new InvalidDataException("This outline-code field requires a leaf value.");
        if (selection && _definition.AllLevelsRequired == true && chain.Count != _masks.Length)
            throw new InvalidDataException("This outline-code field requires every mask level.");
        var pieces = new List<string>(); int length = 0;
        for (int i = 0; i < chain.Count; i++) {
            var current = chain[i];
            if (current.Type.HasValue && current.Type != 21)
                throw new NotSupportedException("Outline-code text evaluation requires text lookup components.");
            string component = current.Value ?? throw new InvalidDataException("An outline-code component has no value.");
            if (component.Length == 0) throw new InvalidDataException("An outline-code component cannot be empty.");
            if (_masks.Length > 0) {
                if (i >= _masks.Length) throw new InvalidDataException("The outline-code hierarchy exceeds its masks.");
                var mask = _masks[i];
                if (mask.Length > 0 && component.Length != mask.Length ||
                    component.Any(c => mask.Type == 0 ? c < '0' || c > '9' : mask.Type == 1 ? !char.IsUpper(c) : mask.Type == 2 && !char.IsLower(c)))
                    throw new InvalidDataException("An outline-code component does not match its mask.");
            }
            if (i > 0) pieces.Add(_masks.Length == 0 ? "." : _masks[i - 1].Separator ?? ".");
            pieces.Add(component); length += component.Length + (i > 0 ? pieces[pieces.Count - 2].Length : 0);
            if (length > 4096) throw new InvalidDataException("The outline-code text exceeds 4096 characters.");
        }
        return string.Concat(pieces);
    }
    private List<ProjectOutlineCodeLookupValue> Chain(ProjectOutlineCodeLookupValue value) {
        var chain = new List<ProjectOutlineCodeLookupValue>(); var visited = new HashSet<int>();
        while (true) {
            if (!value.ValueId.HasValue || !_values.TryGetValue(value.ValueId.Value, out var member) || member != value)
                throw new ArgumentException("The value does not belong to this outline-code table.");
            if (chain.Count >= 64 || !visited.Add(value.ValueId.Value))
                throw new InvalidDataException("The outline-code hierarchy has a cycle or exceeds 64 levels.");
            chain.Add(value);
            if ((value.ParentValueId ?? 0) == 0) break;
            value = _values[value.ParentValueId!.Value];
        }
        chain.Reverse(); return chain;
    }
    internal static bool SameGuid(string? first, string? second) =>
        Guid.TryParse(first, out var a) && Guid.TryParse(second, out var b) && a == b;
}
