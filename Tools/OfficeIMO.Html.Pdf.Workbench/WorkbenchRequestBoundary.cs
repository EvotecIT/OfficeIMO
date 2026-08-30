using Microsoft.Extensions.Primitives;

namespace OfficeIMO.Html.Pdf.Workbench;

internal static class WorkbenchRequestBoundary {
    internal static bool IsAllowedHost(HostString requestHost, Uri listenUri) {
        if (!requestHost.HasValue
            || !string.Equals(NormalizeHost(requestHost.Host), NormalizeHost(listenUri.Host), StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        int? expectedPort = listenUri.IsDefaultPort ? null : listenUri.Port;
        return requestHost.Port == expectedPort;
    }

    internal static bool IsAllowedWebSocketOrigin(StringValues origins, Uri listenUri) {
        if (origins.Count != 1
            || !Uri.TryCreate(origins[0], UriKind.Absolute, out Uri? origin)
            || !string.IsNullOrEmpty(origin.UserInfo)
            || origin.AbsolutePath != "/"
            || !string.IsNullOrEmpty(origin.Query)
            || !string.IsNullOrEmpty(origin.Fragment)) {
            return false;
        }

        return string.Equals(origin.Scheme, listenUri.Scheme, StringComparison.OrdinalIgnoreCase)
               && string.Equals(NormalizeHost(origin.Host), NormalizeHost(listenUri.Host), StringComparison.OrdinalIgnoreCase)
               && origin.Port == listenUri.Port;
    }

    private static string NormalizeHost(string host) => host.Trim('[', ']');
}
