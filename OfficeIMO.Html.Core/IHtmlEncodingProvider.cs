using System.Text;

namespace OfficeIMO.Html.Dom;

/// <summary>Resolves web encoding labels without exposing a charset-library type. Providers must be safe to share.</summary>
public interface IHtmlEncodingProvider {
    /// <summary>Returns an encoding for a recognized HTML label, or null when unsupported.</summary>
    Encoding? ResolveLabel(string label);
    /// <summary>Resolves a charset declaration in an HTML meta content value, or null when absent or unsupported.</summary>
    Encoding? ResolveMetaContent(string content);
}
