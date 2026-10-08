using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Access {
    /// <summary>Whether native object metadata has been decoded.</summary>
    public enum AccessCatalogStatus {
        /// <summary>New model; its collections contain the modeled objects.</summary>
        Modeled,
        /// <summary>The header is recognized, but object catalogs have not been decoded.</summary>
        NotDecoded,
        /// <summary>Native catalog and selected table schemas are decoded; application payloads have separate qualification.</summary>
        Decoded
    }

    /// <summary>Bounded, inert header evidence. It does not certify that the database is valid or unprotected.</summary>
    public sealed class AccessInspection {
        internal AccessInspection(AccessFileFormat format, AccessFormatProfile profile, int version, int subVersion, int pageSize, long length, string hash) {
            Format = format; Profile = profile; HeaderVersion = version; HeaderSubVersion = subVersion; PageSize = pageSize; Length = length; Sha256 = hash;
            Diagnostics = Array.AsReadOnly(new[] {
                new AccessDiagnostic("access.catalog.not-decoded", "Header inspection does not decode catalog, tables, queries, forms, reports, macros or VBA."),
                new AccessDiagnostic("access.protection.not-assessed", "Password, encryption and signature carriers have not been assessed. The header alone cannot establish protection state."),
                new AccessDiagnostic("access.structure.not-validated", "Page alignment is checked; catalog, allocation and page-chain integrity are not yet qualified.")
            });
        }
        /// <summary>Detected family, independent of filename.</summary>
        public AccessFileFormat Format { get; }
        /// <summary>Detected generation, including unqualified generations.</summary>
        public AccessFormatProfile Profile { get; }
        /// <summary>Raw header generation code.</summary>
        public int HeaderVersion { get; }
        /// <summary>Raw compatibility subversion; its feature flags are not decoded by this slice.</summary>
        public int HeaderSubVersion { get; }
        /// <summary>Physical page size for the recognized generation.</summary>
        public int PageSize { get; }
        /// <summary>Complete snapshot byte length.</summary>
        public long Length { get; }
        /// <summary>Physical page count, checked against input limits.</summary>
        public long PageCount => Length / PageSize;
        /// <summary>SHA-256 of the full bounded snapshot; no credentials or active content are interpreted.</summary>
        public string Sha256 { get; }
        /// <summary>Exact boundaries of inspection.</summary>
        public IReadOnlyList<AccessDiagnostic> Diagnostics { get; }

        internal static AccessInspection Read(byte[] bytes, AccessLoadOptions options, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            if (bytes.Length < 21 || bytes[0] != 0 || bytes[1] != 1 || bytes[2] != 0 || bytes[3] != 0 || bytes[19] != 0)
                throw new InvalidDataException("Not a recognized Access database header.");
            string engine = Encoding.ASCII.GetString(bytes, 4, 15);
            bool jet = engine == "Standard Jet DB";
            if (!jet && engine != "Standard ACE DB") throw new InvalidDataException("Unknown Access database engine signature.");
            int version = bytes[20];
            AccessFormatProfile profile = (jet, version) switch {
                (true, 0) => AccessFormatProfile.Jet3, (true, 1) => AccessFormatProfile.Jet4,
                (false, 2) => AccessFormatProfile.Ace12, (false, 3) => AccessFormatProfile.Ace14,
                (false, 5) => AccessFormatProfile.Ace16, (false, 6) => AccessFormatProfile.Ace17,
                _ => AccessFormatProfile.Unknown
            };
            if (profile == AccessFormatProfile.Unknown) throw new NotSupportedException($"Access header version {version} for {engine} is unqualified; page size cannot be assumed.");
            int pageSize = profile == AccessFormatProfile.Jet3 ? 2048 : 4096;
            if (bytes.Length < pageSize * 3 || bytes.Length % pageSize != 0) throw new InvalidDataException("Access input is truncated or is not aligned to its generation's physical pages.");
            if ((long)bytes.Length / pageSize > options.MaxPages) throw new InvalidDataException("Access input exceeds MaxPages.");
            // Hash incrementally so cancellation remains observable even at the snapshot limit.
            using SHA256 hash = SHA256.Create();
            for (int offset = 0; offset < bytes.Length;) {
                cancellationToken.ThrowIfCancellationRequested();
                int count = Math.Min(81920, bytes.Length - offset);
                hash.TransformBlock(bytes, offset, count, bytes, offset); offset = checked(offset + count);
            }
            hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            string digest = BitConverter.ToString(hash.Hash!).Replace("-", "").ToLowerInvariant();
            cancellationToken.ThrowIfCancellationRequested();
            return new AccessInspection(jet ? AccessFileFormat.Mdb : AccessFileFormat.Accdb, profile, version, bytes[21], pageSize, bytes.Length, digest);
        }
    }
}
