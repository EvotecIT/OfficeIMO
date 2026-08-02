using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>Controls resource limits and package inspection when loading Visio documents into memory.</summary>
    public sealed class VisioLoadOptions {
        /// <summary>Default maximum VSDX bytes buffered while loading: 512 MiB.</summary>
        public const long DefaultMaxInputBytes = 512L * 1024L * 1024L;

        /// <summary>
        /// Maximum VSDX bytes buffered by stream and asynchronous load APIs. Default: 512 MiB. Set to null to disable this compatibility guard.
        /// </summary>
        public long? MaxInputBytes { get; set; } = DefaultMaxInputBytes;

        /// <summary>Optional Office package resource limits and active-content policies.</summary>
        public OfficePackageSecurityOptions? PackageSecurity { get; set; }
    }
}
