using System;
using System.Collections.Generic;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Pdf;

/// <summary>PowerShell-captured column composition, delegated to the native row builder.</summary>
internal sealed class PdfColumnContent {
    internal PdfColumnWidth Width { get; }
    internal List<Action<PdfContentBuilder>> Actions { get; } = new();
    internal PdfColumnContent(PdfColumnWidth width) => Width = width;
    internal void Apply(PdfContentBuilder content) { foreach (var action in Actions) action(content); }
}
