using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace OfficeIMO.Provenance;

/// <summary>An immutable text review whose selections are bound to the exact reviewed source.</summary>
public sealed class OfficeTextIntegrityReview {
    private readonly Encoding _encoding;
    private readonly byte[] _preamble;
    private OfficeTextIntegrityReview(string text, Encoding encoding, byte[] preamble, OfficeTextIntegrityOptions options,
        string location, CancellationToken cancellationToken, string? inputSha256 = null) {
        Text = text; _encoding = encoding; _preamble = preamble; InputSha256 = inputSha256;
        Report = OfficeTextIntegrityInspector.Inspect(text, options, location, cancellationToken);
        var codeUnits = new byte[checked(text.Length * 2)];
        for (int index = 0; index < text.Length; index++) { codeUnits[index * 2] = (byte)text[index]; codeUnits[index * 2 + 1] = (byte)(text[index] >> 8); }
        TextSha256 = ComputeSha256(codeUnits);
    }
    /// <summary>Exact decoded source; offsets are UTF-16 code units in this string.</summary>
    public string Text { get; }
    /// <summary>SHA-256 of the exact UTF-16 little-endian code units, without a BOM.</summary>
    public string TextSha256 { get; }
    /// <summary>SHA-256 of the original encoded bytes for a file review; null for pasted text.</summary>
    public string? InputSha256 { get; }
    /// <summary>Encoding used when exporting the separate copy.</summary>
    public string EncodingName => _encoding.WebName;
    /// <summary>Whether a source byte-order mark is retained on export.</summary>
    public bool HasByteOrderMark => _preamble.Length != 0;
    /// <summary>Bounded Unicode evidence requiring contextual review.</summary>
    public OfficeTextIntegrityReport Report { get; }
    /// <summary>Reviews pasted text. Export uses UTF-8 without a byte-order mark.</summary>
    public static OfficeTextIntegrityReview Inspect(string text, OfficeTextIntegrityOptions? options = null,
        string location = "Text", CancellationToken cancellationToken = default) => new(
            text ?? throw new ArgumentNullException(nameof(text)), new UTF8Encoding(false, true), Array.Empty<byte>(),
            options ?? new OfficeTextIntegrityOptions(), location, cancellationToken);
    /// <summary>Reviews strict UTF-8 or BOM-declared UTF-16/32, preserving encoding and the original BOM on export.</summary>
    public static OfficeTextIntegrityReview Inspect(byte[] data, OfficeTextIntegrityOptions? options = null,
        string location = "Text", CancellationToken cancellationToken = default) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        options ??= new OfficeTextIntegrityOptions();
        if (data.LongLength > options.MaxEncodedBytes) throw new InvalidDataException("The text exceeds the encoded byte limit.");
        data = (byte[])data.Clone();
        string text = OfficeTextIntegrityInspector.DecodeText(data, options.MaxCharacters, null, cancellationToken,
            out Encoding encoding, out int preambleLength);
        return new OfficeTextIntegrityReview(text, encoding, data.Take(preambleLength).ToArray(), options, location, cancellationToken, ComputeSha256(data));
    }
    /// <summary>Removes only selected finding indices after confirming the source still matches. No normalization is applied.</summary>
    public string RemoveSelected(string currentText, IEnumerable<int> selectedIndices) {
        if (!string.Equals(currentText, Text, StringComparison.Ordinal))
            throw new InvalidOperationException("The text changed after inspection. Inspect it again before applying selections.");
        if (selectedIndices == null) throw new ArgumentNullException(nameof(selectedIndices));
        int[] indices = selectedIndices.Distinct().OrderBy(index => index).ToArray();
        if (indices.Any(index => index < 0 || index >= Report.Findings.Count)) throw new ArgumentOutOfRangeException(nameof(selectedIndices));
        return OfficeTextIntegrityCleaner.RemoveSelected(Text, indices.Select(index => Report.Findings[index]));
    }
    /// <summary>Calculates a lowercase SHA-256 digest of encoded bytes.</summary>
    public static string ComputeSha256(byte[] data) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        using (SHA256 algorithm = SHA256.Create()) return BitConverter.ToString(algorithm.ComputeHash(data)).Replace("-", "").ToLowerInvariant();
    }
    /// <summary>Exports a separate copy in the original encoding, preserving line endings and unselected code points.</summary>
    public byte[] ExportSelected(string currentText, IEnumerable<int> selectedIndices) {
        byte[] content = _encoding.GetBytes(RemoveSelected(currentText, selectedIndices));
        var output = new byte[_preamble.Length + content.Length];
        Buffer.BlockCopy(_preamble, 0, output, 0, _preamble.Length);
        Buffer.BlockCopy(content, 0, output, _preamble.Length, content.Length);
        return output;
    }
}
