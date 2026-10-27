using OfficeIMO.Internal;
using System.Security.Cryptography;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private sealed class CapturedSourceGuard(string path, byte[] expectedHash, long maximumBytes,
        OfficeConversionExecutionInput? sourceOverride) : IOfficePublicationGuard {
        public async Task<bool> CanPublishAsync(string destination, bool isDirectory, CancellationToken cancellationToken) {
            using Stream current = sourceOverride == null ? OfficeConversionInputReader.OpenLocalRead(path)
                : await sourceOverride.OpenRead(cancellationToken).ConfigureAwait(false);
            byte[] actual = await ComputeSourceHashAsync(current, maximumBytes, cancellationToken).ConfigureAwait(false);
            if (!OfficeRuntimeCompatibility.FixedTimeEquals(expectedHash, actual))
                throw new IOException("The source changed during conversion; no output was published.");
            return true;
        }
    }
}
