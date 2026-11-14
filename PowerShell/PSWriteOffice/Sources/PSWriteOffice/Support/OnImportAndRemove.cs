using System.Management.Automation;
#if FRAMEWORK
using System.Runtime.InteropServices;
#endif

/// <summary>
/// Registers PSWriteOffice's Windows PowerShell assembly resolver when the module is imported and removes it when the module is removed.
/// </summary>
public class OnModuleImportAndRemove : IModuleAssemblyInitializer, IModuleAssemblyCleanup {
    /// <summary>
    /// OnImport is called when the module is imported.
    /// </summary>
    public void OnImport() {
#if FRAMEWORK
        if (IsNetFramework()) {
            DesktopAssemblyResolver.Register();
        }
#endif
    }

    /// <summary>
    /// OnRemove is called when the module is removed.
    /// </summary>
    /// <param name="module">The module being removed.</param>
    public void OnRemove(PSModuleInfo module) {
#if FRAMEWORK
        if (IsNetFramework()) {
            DesktopAssemblyResolver.Unregister();
        }
#endif
    }

#if FRAMEWORK
    private static bool IsNetFramework() {
        return RuntimeInformation.FrameworkDescription.StartsWith(".NET Framework", StringComparison.OrdinalIgnoreCase);
    }
#endif
}
