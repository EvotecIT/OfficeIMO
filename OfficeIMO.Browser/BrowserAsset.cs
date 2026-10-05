using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Browser;

/// <summary>A browser script whose content and immutable file name come from the same embedded resource.</summary>
public sealed class BrowserAsset {
    private static readonly UTF8Encoding Encoding = new UTF8Encoding(false);
    private readonly Lazy<string> _content;
    private readonly Lazy<string> _hash;

    internal BrowserAsset(string fileName) {
        FileName = fileName;
        _content = new Lazy<string>(() => {
            using Stream stream = typeof(BrowserAsset).Assembly.GetManifestResourceStream("OfficeIMO.Browser.Assets." + fileName)
                ?? throw new InvalidOperationException("Missing embedded browser asset: " + fileName);
            using var reader = new StreamReader(stream, Encoding, detectEncodingFromByteOrderMarks: false);
            return reader.ReadToEnd();
        });
        _hash = new Lazy<string>(() => {
            using var sha = SHA256.Create();
            byte[] hash = sha.ComputeHash(Encoding.GetBytes(Content));
            var text = new StringBuilder(16);
            for (int i = 0; i < 8; i++) text.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
            return text.ToString();
        });
    }

    /// <summary>The plain file name, for hosts that provide their own asset pipeline.</summary>
    public string FileName { get; }

    /// <summary>The complete script, suitable for embedding in a portable page.</summary>
    public string Content => _content.Value;

    /// <summary>The first 16 lowercase hex digits of the SHA-256 of the UTF-8 bytes, without BOM.</summary>
    public string ContentHash => _hash.Value;

    /// <summary>A file name bound to these exact bytes, such as officeimo-xlsx.0123456789abcdef.js.</summary>
    public string HashedFileName => Path.GetFileNameWithoutExtension(FileName) + "." + ContentHash + Path.GetExtension(FileName);
}
