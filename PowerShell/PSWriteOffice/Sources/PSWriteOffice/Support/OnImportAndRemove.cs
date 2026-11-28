using System.Management.Automation;
#if FRAMEWORK
using System.Runtime.InteropServices;
#endif

/// <summary>
/// Registers PSWriteOffice's retired-command hints and its Windows PowerShell assembly resolver when the module is imported,
/// and removes them when the module is removed.
/// </summary>
public class OnModuleImportAndRemove : IModuleAssemblyInitializer, IModuleAssemblyCleanup {
    /// <summary>
    /// OnImport is called when the module is imported.
    /// </summary>
    public void OnImport() {
        PSWriteOffice.Support.OfficeCommandHints.Register();
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
        PSWriteOffice.Support.OfficeCommandHints.Unregister();
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
