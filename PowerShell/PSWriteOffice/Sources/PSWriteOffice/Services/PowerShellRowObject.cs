using System.Linq.Expressions;
using System.Management.Automation;
using System.Reflection;

namespace PSWriteOffice.Services;

/// <summary>Creates the <see cref="PSObject"/> rows that reader cmdlets emit.</summary>
internal static class PowerShellRowObject
{
    // PowerShell 7+ exposes PSObject(int) for initial member capacity, but PowerShellStandard also exposes
    // PSObject(object). Use reflection so older runtimes fall back without wrapping the integer as BaseObject.
    private static readonly Func<int, PSObject>? CapacityFactory = CreateCapacityFactory();

    /// <summary>Creates an empty row object sized for <paramref name="capacity"/> properties.</summary>
    public static PSObject Create(int capacity)
    {
        return CapacityFactory?.Invoke(capacity) ?? new PSObject();
    }

    /// <summary>
    /// Returns whether row properties with these names can be added with <c>preValidated: true</c>, which skips the
    /// per-property duplicate lookup. Empty or duplicate names keep the validated add and its usual error.
    /// </summary>
    public static bool CanAddPrevalidated(IEnumerable<string?> names)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in names)
        {
            if (string.IsNullOrEmpty(name) || !seen.Add(name!))
            {
                return false;
            }
        }

        return true;
    }

    private static Func<int, PSObject>? CreateCapacityFactory()
    {
        ConstructorInfo? constructor = null;
        foreach (var candidate in typeof(PSObject).GetConstructors(BindingFlags.Public | BindingFlags.Instance))
        {
            var parameters = candidate.GetParameters();
            if (parameters.Length == 1 && parameters[0].ParameterType == typeof(int))
            {
                constructor = candidate;
                break;
            }
        }

        if (constructor is null)
        {
            return null;
        }

        try
        {
            var capacity = Expression.Parameter(typeof(int), "capacity");
            return Expression.Lambda<Func<int, PSObject>>(
                Expression.New(constructor, capacity),
                capacity).Compile();
        }
        catch
        {
            return capacity => (PSObject)constructor.Invoke(new object[] { capacity });
        }
    }
}
