using System.Runtime.CompilerServices;

namespace OfficeIMO.Rtf;

/// <summary>Copies the semantic graph without serialization, ingestion policy, or normalization.</summary>
internal sealed class RtfCloneContext {
    private readonly Dictionary<object, object> _copies = new Dictionary<object, object>(ReferenceComparer.Instance);

    internal void Register(object source, object copy) => _copies.Add(source, copy);

    internal T? Clone<T>(T? source) where T : class {
        if (source == null) return null;
        if (_copies.TryGetValue(source, out object? copy)) return (T)copy;
        if (source is IRtfCloneable model) return (T)model.CloneModel(this);
        if (source is byte[] bytes) {
            var clonedBytes = (byte[])bytes.Clone();
            Register(source, clonedBytes);
            return (T)(object)clonedBytes;
        }
        // These syntax and alternate-content objects expose only immutable state.
        if (source is string || source is Uri || source is RtfFieldCodeSyntax || source is RtfHtmlEncapsulation) return source;
        throw new NotSupportedException($"Semantic cloning is not supported for {source.GetType().FullName}.");
    }

    internal List<T> CloneList<T>(List<T> source) {
        if (_copies.TryGetValue(source, out object? copy)) return (List<T>)copy;
        var cloned = new List<T>(source.Count);
        Register(source, cloned);
        foreach (T item in source) {
            cloned.Add(item == null || typeof(T).IsValueType ? item : (T)Clone((object)item)!);
        }
        return cloned;
    }

    private sealed class ReferenceComparer : IEqualityComparer<object> {
        internal static readonly ReferenceComparer Instance = new ReferenceComparer();
        public new bool Equals(object? left, object? right) => ReferenceEquals(left, right);
        public int GetHashCode(object value) => RuntimeHelpers.GetHashCode(value);
    }
}

internal interface IRtfCloneable {
    object CloneModel(RtfCloneContext context);
}
