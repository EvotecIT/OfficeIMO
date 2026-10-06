using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Editor;

/// <summary>A redaction candidate derived from one immutable workspace snapshot.</summary>
internal sealed record PdfRedactionCandidate(PdfRedactionArea Area, StudioRectangle Bounds, string Description);
