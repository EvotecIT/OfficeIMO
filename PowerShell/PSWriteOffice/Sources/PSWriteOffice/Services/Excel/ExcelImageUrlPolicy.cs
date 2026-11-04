using System;
using System.Management.Automation;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelImageUrlPolicy
{
    internal static void RequireWebUrl(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)
            || string.IsNullOrEmpty(uri.Host))
        {
            throw new PSArgumentException("Image Url must use HTTP or HTTPS. Use Path for local images.", nameof(url));
        }
    }
}
