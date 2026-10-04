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
}
