using System.Text;
using AngleSharp.Text;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Providers;

/// <summary>Web charset aliases backed by AngleSharp and the .NET code-page provider.</summary>
/// <remarks>Initialization registers the .NET code-page provider process-wide, as the existing decoder did.
/// Returned encodings are cloned by OfficeIMO before changing fallbacks.</remarks>
public sealed class AngleSharpEncodingProvider : IHtmlEncodingProvider {
    /// <summary>Shared charset provider.</summary>
    public static AngleSharpEncodingProvider Instance { get; } = new AngleSharpEncodingProvider();
    static AngleSharpEncodingProvider() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    /// <inheritdoc />
    public Encoding? ResolveLabel(string label) {
        if (label == null) throw new ArgumentNullException(nameof(label));
        if (string.Equals(label.Trim(), "x-user-defined", StringComparison.OrdinalIgnoreCase)) return TextEncoding.Resolve("windows-1252");
        return TextEncoding.IsSupported(label) ? TextEncoding.Resolve(label) : null;
    }
    /// <inheritdoc />
    public Encoding? ResolveMetaContent(string content) => TextEncoding.Parse(content ?? throw new ArgumentNullException(nameof(content)));
}
