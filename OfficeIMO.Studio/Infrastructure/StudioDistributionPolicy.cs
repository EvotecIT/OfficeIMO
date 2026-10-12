namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Immutable distribution capabilities selected at build time.</summary>
internal static class StudioDistributionPolicy {
    internal static bool IsMacAppStore {
        get {
#if OFFICEIMO_MAC_APP_STORE
            return true;
#else
            return false;
#endif
        }
    }

    internal static bool ExternalToolsAllowed => !IsMacAppStore && !OperatingSystem.IsIOS();

    /// <summary>
    /// Sandboxed and mobile builds reach user files only through locations chosen in a picker: a typed path or a folder
    /// next to a picked file grants no access there.
    /// </summary>
    internal static bool RequiresPickedLocations => IsMacAppStore || OperatingSystem.IsIOS() || OperatingSystem.IsAndroid();
}
