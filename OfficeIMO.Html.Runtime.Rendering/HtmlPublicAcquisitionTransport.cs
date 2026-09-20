using System.Net;

namespace OfficeIMO.Html.Runtime.Rendering;

// Deterministic evidence seam: policy remains owned by the workflow and broker;
// only DNS resolution and the already-validated socket destination are replaced.
internal sealed class HtmlPublicAcquisitionTransport {
    internal HtmlPublicAcquisitionTransport(
        Func<string, CancellationToken, Task<IPAddress[]>> resolveAddresses,
        Func<IPAddress, int, CancellationToken, ValueTask<Stream>> connect) {
        ResolveAddresses = resolveAddresses ?? throw new ArgumentNullException(nameof(resolveAddresses));
        Connect = connect ?? throw new ArgumentNullException(nameof(connect));
    }

    internal Func<string, CancellationToken, Task<IPAddress[]>> ResolveAddresses { get; }
    internal Func<IPAddress, int, CancellationToken, ValueTask<Stream>> Connect { get; }
}
