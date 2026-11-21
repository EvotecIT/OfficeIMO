using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace OfficeIMO.Data;

/// <summary>
/// Discovers the public instance properties that object mapping treats as columns.
/// </summary>
/// <remarks>
/// Indexers are excluded because they have no single value per row. When a derived type hides an
/// inherited property with <c>new</c>, reflection reports both declarations under one name; only the
/// most derived one is visible to callers of the type, so it is the only one returned.
/// </remarks>
internal static class PublicInstanceProperties {
    internal static PropertyInfo[] Get(
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] Type type) {
        PropertyInfo[] properties = type.GetProperties(BindingFlags.Instance | BindingFlags.Public);
        var result = new List<PropertyInfo>(properties.Length);
        Dictionary<string, int>? positions = null;
        foreach (PropertyInfo property in properties) {
            if (property.GetIndexParameters().Length != 0) continue;
            positions ??= new Dictionary<string, int>(properties.Length, StringComparer.Ordinal);
            if (positions.TryGetValue(property.Name, out int existing)) {
                if (InheritanceDepth(property.DeclaringType) > InheritanceDepth(result[existing].DeclaringType)) {
                    result[existing] = property;
                }
                continue;
            }
            positions.Add(property.Name, result.Count);
            result.Add(property);
        }
        return result.ToArray();
    }

    /// <summary>Finds a mapped property by its exact name without failing on hidden inherited declarations.</summary>
    internal static PropertyInfo? Find(
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] Type type,
        string name) {
        PropertyInfo? match = null;
        foreach (PropertyInfo property in type.GetProperties(BindingFlags.Instance | BindingFlags.Public)) {
            if (!string.Equals(property.Name, name, StringComparison.Ordinal) || property.GetIndexParameters().Length != 0) continue;
            if (match is null || InheritanceDepth(property.DeclaringType) > InheritanceDepth(match.DeclaringType)) match = property;
        }
        return match;
    }

    private static int InheritanceDepth(Type? type) {
        int depth = 0;
        for (Type? current = type; current is not null; current = current.BaseType) depth++;
        return depth;
    }
}
