using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Internal {
    /// <summary>Identity comparison for caller-owned graphs that may contain cycles.</summary>
    internal sealed class OfficeReferenceEqualityComparer<T> : IEqualityComparer<T> where T : class {
        internal static OfficeReferenceEqualityComparer<T> Instance { get; } = new OfficeReferenceEqualityComparer<T>();
        public bool Equals(T? left, T? right) => ReferenceEquals(left, right);
        public int GetHashCode(T value) => RuntimeHelpers.GetHashCode(value);
    }
}
