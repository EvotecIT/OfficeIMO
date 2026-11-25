using System.Security.Cryptography;

namespace OfficeIMO.Web.Fonts.Notebook;

/// <summary>Supplies the pinned fonts used by browser notebook reports.</summary>
internal static class NotebookFontAssets {
    internal const string MathSha256 = "49b22767985171bce579307b21d98f4d35012fa39e316acc4ffcab8644dfb69b";
    internal const string MonospaceSha256 = "f2b83c763e8afd21709333370bed4774337fae82267937e2b5aea7e2fbd922c1";
    private static readonly Lazy<byte[]> MathFont = new(() => Load("NotebookMath.ttf", MathSha256));
    private static readonly Lazy<byte[]> MonospaceFont = new(() => Load("NotebookMonospace.ttf", MonospaceSha256));

    internal static byte[] MathBytes => MathFont.Value;
    internal static byte[] MonospaceBytes => MonospaceFont.Value;

    private static byte[] Load(string name, string sha256) {
        using Stream resource = typeof(NotebookFontAssets).Assembly.GetManifestResourceStream(name)
            ?? throw new InvalidOperationException("A notebook font resource is missing.");
        using var output = new MemoryStream();
        resource.CopyTo(output);
        byte[] bytes = output.ToArray();
        if (!string.Equals(Convert.ToHexStringLower(SHA256.HashData(bytes)), sha256, StringComparison.Ordinal)) {
            throw new InvalidDataException("A notebook font does not match its pinned source.");
        }
        return bytes;
    }
}
