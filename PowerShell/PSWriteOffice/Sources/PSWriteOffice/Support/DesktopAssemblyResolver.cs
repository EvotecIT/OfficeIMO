#if FRAMEWORK
using System.Reflection;

/// <summary>
/// Resolves PSWriteOffice's own dependencies in Windows PowerShell, where every module shares one AppDomain.
/// </summary>
/// <remarks>
/// The CLR asks AssemblyResolve handlers in subscription order and uses the first answer. A module imported
/// earlier may answer every request by simple name with its own older copy (dbatools.library hands out
/// Microsoft.Bcl.AsyncInterfaces 8.0 to System.Text.Json 10), so this handler subscribes first. It answers only
/// requests from assemblies in PSWriteOffice's library folder, or requests without a known requester for an
/// identity those assemblies reference, and only with a copy of the same name, culture and public key token that
/// is not older than requested. Every other request falls through to the next handler unchanged.
/// </remarks>
internal static class DesktopAssemblyResolver {
    private static readonly object Sync = new();
    private static readonly ResolveEventHandler Handler = Resolve;
    private static readonly string LibraryDirectory = GetLibraryDirectory();
    private static bool registered;

    [ThreadStatic]
    private static bool resolving;

    internal static void Register() {
        lock (Sync) {
            if (registered) {
                return;
            }

            SubscribeFirst(Handler);
            registered = true;
        }
    }

    internal static void Unregister() {
        lock (Sync) {
            if (!registered) {
                return;
            }

            AppDomain.CurrentDomain.AssemblyResolve -= Handler;
            registered = false;
        }
    }

    private static void SubscribeFirst(ResolveEventHandler handler) {
        AppDomain domain = AppDomain.CurrentDomain;
        // .NET Framework keeps the subscribers in this field and its add/remove accessors lock the AppDomain.
        FieldInfo? field = typeof(AppDomain).GetField("_AssemblyResolve", BindingFlags.Instance | BindingFlags.NonPublic);
        if (field is null || field.FieldType != typeof(ResolveEventHandler)) {
            domain.AssemblyResolve += handler;
            return;
        }

        lock (domain) {
            var existing = (ResolveEventHandler?)field.GetValue(domain);
            field.SetValue(domain, (ResolveEventHandler)Delegate.Combine(handler, existing));
        }
    }

    private static Assembly? Resolve(object? sender, ResolveEventArgs args) {
        if (resolving || args is null) {
            return null;
        }

        resolving = true;
        try {
            Assembly? requester = args.RequestingAssembly;
            if (requester is null || requester.IsDynamic || string.IsNullOrEmpty(requester.Location)) {
                if (!IsReferencedByModule(args.Name)) {
                    return null;
                }
            } else if (!IsModuleAssembly(requester)) {
                return null;
            }

            var requested = new AssemblyName(args.Name);
            string? name = requested.Name;
            if (string.IsNullOrEmpty(name) || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
                return null;
            }

            string candidate = Path.Combine(LibraryDirectory, name + ".dll");
            if (!File.Exists(candidate) || !IsCompatible(requested, AssemblyName.GetAssemblyName(candidate))) {
                return null;
            }

            return Assembly.LoadFrom(candidate);
        } catch {
            return null;
        } finally {
            resolving = false;
        }
    }

    private static bool IsModuleAssembly(Assembly assembly) {
        if (assembly.IsDynamic || string.IsNullOrEmpty(assembly.Location)) {
            return false;
        }

        return Path.GetFullPath(assembly.Location).StartsWith(LibraryDirectory, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsReferencedByModule(string fullName) {
        foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies()) {
            if (!IsModuleAssembly(assembly)) {
                continue;
            }

            foreach (AssemblyName reference in assembly.GetReferencedAssemblies()) {
                if (string.Equals(reference.FullName, fullName, StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool IsCompatible(AssemblyName requested, AssemblyName candidate) {
        if (!string.Equals(requested.Name, candidate.Name, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(requested.CultureName ?? string.Empty, candidate.CultureName ?? string.Empty, StringComparison.OrdinalIgnoreCase) ||
            (requested.Version is not null && (candidate.Version is null || candidate.Version < requested.Version))) {
            return false;
        }

        byte[] expected = requested.GetPublicKeyToken() ?? Array.Empty<byte>();
        byte[] actual = candidate.GetPublicKeyToken() ?? Array.Empty<byte>();
        // Plain BCL code only: the resolver must not depend on assemblies it may be asked to resolve.
        if (expected.Length != actual.Length) {
            return false;
        }

        for (int i = 0; i < expected.Length; i++) {
            if (expected[i] != actual[i]) {
                return false;
            }
        }

        return true;
    }

    private static string GetLibraryDirectory() {
        string directory = Path.GetDirectoryName(Path.GetFullPath(typeof(DesktopAssemblyResolver).Assembly.Location)) ?? string.Empty;
        return directory.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
    }
}
#endif
