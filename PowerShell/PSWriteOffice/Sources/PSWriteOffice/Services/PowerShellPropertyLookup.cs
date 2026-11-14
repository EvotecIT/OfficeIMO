using System.Linq.Expressions;
using System.Management.Automation;
using System.Reflection;

namespace PSWriteOffice.Services;

/// <summary>
/// Looks up row properties by name with the same result as <c>PSObject.Properties[name]</c>, but faster for
/// <see cref="PSCustomObject"/> rows.
/// </summary>
/// <remarks>
/// PowerShell searches an object's own (instance) members before type-table and adapted members, so a property
/// found among the instance members is exactly what <c>Properties[name]</c> returns. Reading the instance member
/// collection directly skips the integrated collection's per-lookup overhead, which is a large share of exporting
/// <c>[pscustomobject]</c> rows. Anything else, including names that are not instance properties, uses the public lookup.
/// </remarks>
internal static class PowerShellPropertyLookup
{
    private static readonly Func<PSObject, PSMemberInfoCollection<PSMemberInfo>>? InstanceMembersAccessor = CreateInstanceMembersAccessor();

    /// <summary>Gets the instance members of a <see cref="PSCustomObject"/> row when the fast path is available.</summary>
    public static PSMemberInfoCollection<PSMemberInfo>? GetInstanceMembers(PSObject ps)
    {
        if (InstanceMembersAccessor == null || ps.BaseObject is not PSCustomObject)
        {
            return null;
        }

        try
        {
            return InstanceMembersAccessor(ps);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>Returns the property that <c>ps.Properties[name]</c> returns.</summary>
    public static PSPropertyInfo? GetProperty(PSObject ps, PSMemberInfoCollection<PSMemberInfo>? instanceMembers, string name)
    {
        if (instanceMembers != null && instanceMembers[name] is PSPropertyInfo property)
        {
            return property;
        }

        return ps.Properties[name];
    }

    private static Func<PSObject, PSMemberInfoCollection<PSMemberInfo>>? CreateInstanceMembersAccessor()
    {
        try
        {
            var property = typeof(PSObject).GetProperty("InstanceMembers", BindingFlags.Instance | BindingFlags.NonPublic);
            if (property == null ||
                property.GetIndexParameters().Length != 0 ||
                !typeof(PSMemberInfoCollection<PSMemberInfo>).IsAssignableFrom(property.PropertyType))
            {
                return null;
            }

            var target = Expression.Parameter(typeof(PSObject), "target");
            var accessor = Expression.Lambda<Func<PSObject, PSMemberInfoCollection<PSMemberInfo>>>(
                Expression.Convert(Expression.Property(target, property), typeof(PSMemberInfoCollection<PSMemberInfo>)),
                target).Compile();

            return Validate(accessor) ? accessor : null;
        }
        catch
        {
            return null;
        }
    }

    // Use the shortcut only when it agrees with the public lookup on this PowerShell version.
    private static bool Validate(Func<PSObject, PSMemberInfoCollection<PSMemberInfo>> accessor)
    {
        var probe = new PSObject();
        probe.Properties.Add(new PSNoteProperty("Probe", 42));
        return probe.BaseObject is PSCustomObject &&
            accessor(probe)["probe"] is PSPropertyInfo property &&
            ReferenceEquals(property.Value, probe.Properties["Probe"]?.Value);
    }
}
