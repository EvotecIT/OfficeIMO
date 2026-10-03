using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using OfficeIMO.OneNote.Markdown;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Html.EditableEvidence;

internal sealed record EditableExport(
    HtmlConversionReport Report,
    string Artifact,
    string ReopenedHtml,
    int? MarkdownTableRows = null,
    int? MarkdownTableColumns = null,
    int? PictureCount = null,
    int? LinkedPictureCount = null,
    IReadOnlyList<NativeTableEvidence>? NativeTables = null);

internal static class EditableTargetExporter {
    internal static EditableExport SaveAndReopen(HtmlConversionDocument source, string target, string output, string documentName) {
        switch (target) {
            case "word": {
                HtmlToWordResult result = source.ToWordDocumentResult();
                string artifact = Path.Combine(output, documentName + ".docx");
                using (WordDocument document = result.RequireValue()) document.Save(artifact);
                using WordDocument loaded = WordDocument.Load(artifact);
                return new EditableExport(result.Report, artifact, loaded.ToHtmlResult().RequireValue(),
                    NativeTables: EditableNativeTableEvidence.FromWord(loaded));
            }
            case "excel": {
                HtmlToExcelResult result = source.ToExcelDocumentResult(new HtmlToExcelOptions {
                    Mode = HtmlImportMode.Generic,
                    ImportEditableLayoutRegions = true
                });
                string artifact = Path.Combine(output, documentName + ".xlsx");
                using (ExcelDocument workbook = result.RequireValue()) workbook.Save(artifact);
                using ExcelDocument loaded = ExcelDocument.Load(artifact);
                return new EditableExport(result.Report, artifact, loaded.ToHtml(),
                    NativeTables: EditableNativeTableEvidence.FromExcel(loaded));
            }
            case "powerpoint": {
                HtmlToPowerPointResult result = source.ToPowerPointPresentationResult(new HtmlToPowerPointOptions {
                    Mode = HtmlImportMode.Generic,
                    ImportEditableLayoutRegions = true
                });
                string artifact = Path.Combine(output, documentName + ".pptx");
                using (PowerPointPresentation presentation = result.RequireValue()) presentation.Save(artifact);
                using PowerPointPresentation loaded = PowerPointPresentation.Load(artifact);
                return new EditableExport(result.Report, artifact, loaded.ToHtml(),
                    PictureCount: loaded.Slides.Sum(slide => slide.Pictures.Count()),
                    LinkedPictureCount: loaded.Slides.Sum(slide => slide.Pictures.Count(picture => picture.Hyperlink != null)),
                    NativeTables: EditableNativeTableEvidence.FromPowerPoint(loaded));
            }
            case "onenote": {
                HtmlToOneNoteSectionResult result = source.ToOneNoteSectionResult();
                string artifact = Path.Combine(output, documentName + ".one");
                result.RequireValue().Save(artifact);
                OneNoteSection loaded = OneNoteSectionReader.Read(artifact);
                int imageIndex = 0;
                return new EditableExport(result.Report, artifact, loaded.ToHtmlDocument(new OneNoteMarkdownOptions {
                    AssetUriResolver = element => {
                        if (element is not OneNoteImage { Payload: not null } image) return null;
                        string? extension = image.MediaType switch {
                            "image/png" => ".png",
                            "image/jpeg" => ".jpg",
                            "image/gif" => ".gif",
                            "image/webp" => ".webp",
                            _ => null
                        };
                        if (extension == null || image.Payload!.Length > 16 * 1024 * 1024) return null;
                        byte[] bytes = image.Payload.ToArray(16 * 1024 * 1024);
                        string directory = Path.Combine(output, "assets");
                        Directory.CreateDirectory(directory);
                        string fileName = $"image-{++imageIndex:D4}{extension}";
                        File.WriteAllBytes(Path.Combine(directory, fileName), bytes);
                        return "assets/" + fileName;
                    }
                }), NativeTables: EditableNativeTableEvidence.FromOneNote(loaded));
            }
            case "rtf": {
                HtmlToRtfResult result = source.ToRtfDocumentResult();
                string artifact = Path.Combine(output, documentName + ".rtf");
                result.RequireValue().Save(artifact);
                RtfDocument loaded = RtfDocument.Load(artifact);
                return new EditableExport(result.Report, artifact, loaded.ToHtml(new RtfToHtmlOptions {
                    EmbedImagesAsDataUri = true,
                    MaxEmbeddedImageBytes = 16 * 1024 * 1024
                }), NativeTables: EditableNativeTableEvidence.FromRtf(loaded));
            }
            case "markdown": {
                HtmlToMarkdownResult result = source.ToMarkdownDocumentResult();
                string artifact = Path.Combine(output, documentName + ".md");
                result.RequireValue().Save(artifact);
                MarkdownDoc loaded = MarkdownDoc.Load(artifact);
                TableBlock? table = loaded.Blocks.OfType<TableBlock>().FirstOrDefault();
                return new EditableExport(result.Report, artifact, loaded.ToMarkdown(),
                    MarkdownTableRows: table?.Rows.Count,
                    MarkdownTableColumns: table?.Headers.Count,
                    NativeTables: EditableNativeTableEvidence.FromMarkdown(loaded));
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(target), target, "The selected editable target is not supported.");
        }
    }
}
