using System;
using System.Management.Automation;
using System.Text;
using System.Threading;
using OfficeIMO.Epub;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.OpenDocument;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Rtf;
using OfficeIMO.Visio;
using OfficeIMO.Word;

namespace PSWriteOffice.Services;

/// <summary>Captures current native content without saving to or closing a caller-owned document.</summary>
internal sealed class OfficeLiveDocumentSource {
    private readonly Func<CancellationToken, byte[]> _serialize;

    private OfficeLiveDocumentSource(string extension, string? path, Func<CancellationToken, byte[]> serialize) {
        Extension = extension;
        Path = path;
        _serialize = serialize;
    }

    internal string Extension { get; }
    internal string? Path { get; }
    internal string Name => "source" + Extension;

    internal byte[] Capture(CancellationToken token) {
        token.ThrowIfCancellationRequested();
        byte[] bytes = _serialize(token);
        token.ThrowIfCancellationRequested();
        return bytes;
    }

    internal static OfficeLiveDocumentSource Create(object value) {
        while (value is PSObject wrapper && !ReferenceEquals(value, wrapper.BaseObject)) value = wrapper.BaseObject;
        // Password=null is an explicit one-time plaintext snapshot, not a change to the caller's saved security policy.
        return value switch {
            WordDocument word => new(".docx", word.FilePath, _ => word.ToBytes(new WordSaveOptions { Password = null })),
            ExcelDocument excel => new(".xlsx", excel.FilePath, _ => excel.ToBytes(new ExcelSaveOptions { Password = null })),
            PowerPointDocument slides => new(".pptx", slides.FilePath, _ => slides.ToBytes(new PowerPointSaveOptions { Password = null })),
            PdfDocument pdf => new(".pdf", pdf.FilePath, token => pdf.ToBytes(cancellationToken: token)),
            MarkdownDocument markdown => new(".md", markdown.FilePath, token => markdown.ToBytes(cancellationToken: token)),
            HtmlConversionDocument html => new(".html", null, _ => Encoding.UTF8.GetBytes(html.ExportSourceHtml())),
            VisioDocument visio => new(".vsdx", visio.FilePath, _ => visio.ToBytes()),
            RtfDocument rtf => new(".rtf", null, _ => rtf.ToBytes()),
            OdfDocument odf => new(OdfExtension(odf.Kind), odf.FilePath, _ => odf.ToBytes()),
            EpubPublication epub => new(".epub", null, token => {
                var result = epub.Write(cancellationToken: token);
                result.Report.RequireNoLoss();
                return result.Bytes;
            }),
            _ => throw new PSArgumentException("Document requires a live Word, Excel, PowerPoint, PDF, Markdown, HTML, Visio, RTF, OpenDocument or EPUB document.", nameof(value))
        };
    }

    private static string OdfExtension(OdfDocumentKind kind) => kind switch {
        OdfDocumentKind.Text => ".odt",
        OdfDocumentKind.Spreadsheet => ".ods",
        OdfDocumentKind.Presentation => ".odp",
        OdfDocumentKind.Graphics => ".odg",
        _ => throw new PSArgumentException("Choose a supported OpenDocument kind.")
    };
}
