using OfficeIMO.Security;

namespace OfficeIMO.Word {
    /// <summary>Trust, revocation, timestamp, and resource policy for OPC XML-signature validation.</summary>
    public sealed class WordSignatureValidationOptions {
        /// <summary>Gets signer certificate-chain and revocation policy.</summary>
        public CertificateValidationOptions CertificateValidation { get; } = new CertificateValidationOptions();

        /// <summary>Gets timestamp-authority certificate-chain and revocation policy.</summary>
        public CertificateValidationOptions TimestampCertificateValidation { get; } = new CertificateValidationOptions();

        /// <summary>Gets or sets whether XML DSig signature-value and signed-object validation is performed.</summary>
        public bool ValidateCryptographicSignature { get; set; } = true;

        /// <summary>Gets or sets whether embedded RFC 3161 timestamp tokens are validated.</summary>
        public bool ValidateTimestamps { get; set; } = true;

        /// <summary>Gets or sets the maximum number of XML signature parts. Defaults to 32.</summary>
        public int MaxSignatureParts { get; set; } = 32;

        /// <summary>Gets or sets the maximum encoded OPC package bytes accepted. Defaults to 512 MiB.</summary>
        public long MaxPackageBytes { get; set; } = 512L * 1024 * 1024;

        /// <summary>Gets or sets the maximum number of OPC package entries accepted. Defaults to 10,000.</summary>
        public int MaxPackageParts { get; set; } = 10000;

        /// <summary>Gets or sets the maximum uncompressed bytes read from one signed package part. Defaults to 256 MiB.</summary>
        public long MaxPartBytes { get; set; } = 256L * 1024 * 1024;

        /// <summary>Gets or sets the maximum authenticated XML signature references per signature part. Defaults to 4,096.</summary>
        public int MaxSignedReferences { get; set; } = 4096;

        /// <summary>Gets or sets the maximum aggregate uncompressed package-part bytes digested per signature part. Defaults to 512 MiB.</summary>
        public long MaxTotalDigestBytes { get; set; } = 512L * 1024 * 1024;

        /// <summary>Gets or sets the maximum bytes read from one signature part. Defaults to 16 MiB.</summary>
        public long MaxSignatureBytes { get; set; } = 16L * 1024 * 1024;

        /// <summary>Gets or sets the maximum embedded or related certificates per signature. Defaults to 64.</summary>
        public int MaxCertificates { get; set; } = 64;

        /// <summary>Gets or sets the maximum encoded bytes read from one related signer certificate. Defaults to 4 MiB.</summary>
        public long MaxCertificateBytes { get; set; } = 4L * 1024 * 1024;

        /// <summary>Gets or sets the maximum aggregate certificate bytes decoded or read across all signature parts. Defaults to 64 MiB.</summary>
        public long MaxTotalCertificateBytes { get; set; } = 64L * 1024 * 1024;

        /// <summary>Gets or sets the maximum RFC 3161 timestamp-token count per signature. Defaults to 16.</summary>
        public int MaxTimestampTokens { get; set; } = 16;

        /// <summary>Gets or sets the maximum encoded bytes per RFC 3161 timestamp token. Defaults to 16 MiB.</summary>
        public long MaxTimestampBytes { get; set; } = 16L * 1024 * 1024;
    }
}
