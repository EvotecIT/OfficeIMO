using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>Controls resource limits and package inspection when loading Visio documents into memory.</summary>
    public sealed class VisioLoadOptions {
        /// <summary>Creates bounded options that reject active, embedded, and externally linked package content.</summary>
        public static VisioLoadOptions UntrustedDefaults => new VisioLoadOptions {
            PackageSecurity = OfficePackageSecurityOptions.UntrustedDefaults
        };

        /// <summary>Resource budgets applied while buffering and parsing Visio sources.</summary>
        public VisioLoadLimits Limits { get; set; } = new VisioLoadLimits();

        /// <summary>Optional Office package resource limits and active-content policies.</summary>
        public OfficePackageSecurityOptions? PackageSecurity { get; set; }
    }

    /// <summary>Resource budgets applied while loading Visio documents. Limits must be positive; a nullable limit set to null is not enforced.</summary>
    public sealed class VisioLoadLimits {
        /// <summary>Default maximum VSDX bytes buffered while loading: 512 MiB.</summary>
        public const long DefaultMaxInputBytes = 512L * 1024L * 1024L;

        /// <summary>
        /// Maximum VSDX bytes buffered by stream and asynchronous load APIs. Default: 512 MiB. Set to null to disable this guard.
        /// </summary>
        public long? MaxInputBytes { get; set; } = DefaultMaxInputBytes;

        /// <summary>Maximum characters in a legacy XML document. Default: 30 million.</summary>
        public long MaxLegacyXmlCharacters { get; set; } = 30_000_000;
        /// <summary>Maximum characters in each VSDX content XML part. Default: 30 million.</summary>
        public long MaxPackageXmlCharacters { get; set; } = 30_000_000;
        /// <summary>Maximum legacy XML nesting depth, enforced before materialization.</summary>
        public int MaxLegacyXmlDepth { get; set; } = 128;
        /// <summary>Maximum legacy XML elements and cumulative attributes, enforced before materialization.</summary>
        public int MaxLegacyXmlElements { get; set; } = 1_000_000;
    }
}
