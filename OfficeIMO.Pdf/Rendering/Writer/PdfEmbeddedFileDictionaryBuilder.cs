using System.Globalization;
using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

internal static class PdfEmbeddedFileDictionaryBuilder {
    internal static string BuildEmbeddedFileStreamDictionary(PdfEmbeddedFile file, byte[] data) {
        Guard.NotNull(file, nameof(file));
        Guard.NotNullOrEmpty(data, nameof(data));
        int length = data.Length;

        var sb = new StringBuilder();
        sb.Append("<< /Type /EmbeddedFile");
        if (!string.IsNullOrWhiteSpace(file.MimeType)) {
            sb.Append(" /Subtype /")
                .Append(PdfSyntaxEscaper.Name(file.MimeType!));
        }

        sb.Append(" /Length ")
            .Append(length.ToString(CultureInfo.InvariantCulture))
            .Append(" /Params << /Size ")
            .Append(length.ToString(CultureInfo.InvariantCulture))
            .Append(" /CheckSum ")
            .Append(PdfSyntaxEscaper.HexString(ComputePdfEmbeddedFileChecksum(data)))
            .Append(" >>")
            .Append(" >>");
        return sb.ToString();
    }

    internal static string BuildFileSpecificationObject(PdfEmbeddedFile file, int embeddedFileObjectId) {
        Guard.NotNull(file, nameof(file));

        var sb = new StringBuilder();
        sb.Append("<< /Type /Filespec /F ")
            .Append(PdfSyntaxEscaper.TextString(file.FileName))
            .Append(" /UF ")
            .Append(PdfSyntaxEscaper.TextString(file.FileName))
            .Append(" /EF << /F ")
            .Append(PdfSyntaxEscaper.IndirectReference(embeddedFileObjectId))
            .Append(" /UF ")
            .Append(PdfSyntaxEscaper.IndirectReference(embeddedFileObjectId))
            .Append(" >> /AFRelationship /")
            .Append(GetRelationshipName(file.Relationship));

        if (file.Description != null) {
            sb.Append(" /Desc ")
                .Append(PdfSyntaxEscaper.TextString(file.Description));
        }

        sb.Append(" >>\n");
        return sb.ToString();
    }

    internal static string BuildEmbeddedFilesNameTree(IReadOnlyList<(string FileName, int FileSpecId)> files) {
        Guard.NotNull(files, nameof(files));
        if (files.Count == 0) {
            throw new ArgumentException("PDF embedded files name tree requires at least one file.", nameof(files));
        }

        var fileNames = new HashSet<string>(StringComparer.Ordinal);
        var sortedFiles = new List<(string FileName, int FileSpecId)>(files.Count);
        foreach ((string fileName, int fileSpecId) in files) {
            Guard.NotNullOrWhiteSpace(fileName, nameof(files));
            if (!fileNames.Add(fileName)) {
                throw new ArgumentException("PDF embedded files name tree names must be unique.", nameof(files));
            }

            sortedFiles.Add((fileName, fileSpecId));
        }

        sortedFiles.Sort((left, right) => StringComparer.Ordinal.Compare(left.FileName, right.FileName));

        var sb = new StringBuilder();
        sb.Append("<< /Names [");
        foreach ((string fileName, int fileSpecId) in sortedFiles) {
            sb.Append(PdfSyntaxEscaper.TextString(fileName))
                .Append(' ')
                .Append(PdfSyntaxEscaper.IndirectReference(fileSpecId))
                .Append(' ');
        }

        if (sb[sb.Length - 1] == ' ') {
            sb.Length--;
        }

        sb.Append("] >>\n");
        return sb.ToString();
    }

    internal static string GetRelationshipName(PdfAssociatedFileRelationship relationship) {
        Guard.AssociatedFileRelationship(relationship, nameof(relationship));
        return relationship switch {
            PdfAssociatedFileRelationship.Source => "Source",
            PdfAssociatedFileRelationship.Data => "Data",
            PdfAssociatedFileRelationship.Alternative => "Alternative",
            PdfAssociatedFileRelationship.Supplement => "Supplement",
            _ => "Unspecified"
        };
    }

    #pragma warning disable CA5351, CA1850 // PDF embedded-file /CheckSum is a deterministic file-identification checksum, not a security boundary; use Create for net472.
    private static byte[] ComputePdfEmbeddedFileChecksum(byte[] data) {
        using (MD5 md5 = MD5.Create()) {
            return md5.ComputeHash(data);
        }
    }
    #pragma warning restore CA5351, CA1850
}
