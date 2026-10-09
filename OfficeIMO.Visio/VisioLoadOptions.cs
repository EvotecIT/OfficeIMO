using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>Controls resource limits and package inspection when loading Visio documents into memory.</summary>
    public sealed class VisioLoadOptions {
        /// <summary>Default maximum VSDX bytes buffered while loading: 512 MiB.</summary>
        public const long DefaultMaxInputBytes = 512L * 1024L * 1024L;

        /// <summary>Creates bounded options that reject active, embedded, and externally linked package content.</summary>
        public static VisioLoadOptions UntrustedDefaults => new VisioLoadOptions {
            PackageSecurity = OfficePackageSecurityOptions.UntrustedDefaults
        };

        /// <summary>
        /// Maximum VSDX bytes buffered by stream and asynchronous load APIs. Default: 512 MiB. Set to null to disable this compatibility guard.
        /// </summary>
        public long? MaxInputBytes { get; set; } = DefaultMaxInputBytes;

        /// <summary>Maximum characters in a legacy XML document. Default: 30 million.</summary>
        public long MaxLegacyXmlCharacters { get; set; } = 30_000_000;
        /// <summary>Maximum legacy XML nesting depth, enforced before materialization.</summary>
        public int MaxLegacyXmlDepth { get; set; } = 128;
        /// <summary>Maximum legacy XML elements and cumulative attributes, enforced before materialization.</summary>
        public int MaxLegacyXmlElements { get; set; } = 1_000_000;

        /// <summary>Optional Office package resource limits and active-content policies.</summary>
        public OfficePackageSecurityOptions? PackageSecurity { get; set; }
    }
}
