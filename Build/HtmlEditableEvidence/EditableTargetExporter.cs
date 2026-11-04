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
    OfficeConversionReport Report,
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
                string artifact = Path.Combine(output, documentName + ".docx");
                OfficeConversionReport report = source.SaveAsWord(artifact);
                using WordDocument loaded = WordDocument.Load(artifact);
                return new EditableExport(report, artifact, loaded.ToHtml(),
                    NativeTables: EditableNativeTableEvidence.FromWord(loaded));
            }
            case "excel": {
                string artifact = Path.Combine(output, documentName + ".xlsx");
                OfficeConversionReport report = source.SaveAsExcel(artifact, new HtmlExcelExportOptions {
                    Mode = HtmlImportMode.Generic,
                    ImportEditableLayoutRegions = true
                });
                using ExcelDocument loaded = ExcelDocument.Load(artifact);
                return new EditableExport(report, artifact, loaded.ToHtml(),
                    NativeTables: EditableNativeTableEvidence.FromExcel(loaded));
            }
            case "powerpoint": {
                string artifact = Path.Combine(output, documentName + ".pptx");
                OfficeConversionReport report = source.SaveAsPowerPoint(artifact, new HtmlPowerPointExportOptions {
                    Mode = HtmlImportMode.Generic,
                    ImportEditableLayoutRegions = true
                });
                using PowerPointDocument loaded = PowerPointDocument.Load(artifact);
                return new EditableExport(report, artifact, loaded.ToHtml(),
                    PictureCount: loaded.Slides.Sum(slide => slide.Images.Count()),
                    LinkedPictureCount: loaded.Slides.Sum(slide => slide.Images.Count(picture => picture.Hyperlink != null)),
                    NativeTables: EditableNativeTableEvidence.FromPowerPoint(loaded));
            }
            case "onenote": {
                string artifact = Path.Combine(output, documentName + ".one");
                source.ToOneNoteSection(null, out OfficeConversionReport report).Save(artifact);
                OneNoteSection loaded = OneNoteSectionReader.Read(artifact);
                int imageIndex = 0;
                return new EditableExport(report, artifact, loaded.ToHtml(new OfficeIMO.OneNote.Html.OneNoteHtmlExportOptions { ProjectionOptions = new OneNoteMarkdownExportOptions {
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
                } }), NativeTables: EditableNativeTableEvidence.FromOneNote(loaded));
            }
            case "rtf": {
                string artifact = Path.Combine(output, documentName + ".rtf");
                OfficeConversionReport report = source.SaveAsRtf(artifact);
                RtfDocument loaded = RtfDocument.Load(artifact);
                return new EditableExport(report, artifact, loaded.ToHtml(new RtfHtmlExportOptions {
                    EmbedImagesAsDataUri = true,
                    MaxEmbeddedImageBytes = 16 * 1024 * 1024
                }), NativeTables: EditableNativeTableEvidence.FromRtf(loaded));
            }
            case "markdown": {
                string artifact = Path.Combine(output, documentName + ".md");
                OfficeConversionReport report = source.SaveAsMarkdown(artifact);
                using MarkdownDocument loaded = MarkdownDocument.Load(artifact);
                TableBlock? table = loaded.Blocks.OfType<TableBlock>().FirstOrDefault();
                return new EditableExport(report, artifact, loaded.ToMarkdown(),
                    MarkdownTableRows: table?.Rows.Count,
                    MarkdownTableColumns: table?.Headers.Count,
                    NativeTables: EditableNativeTableEvidence.FromMarkdown(loaded));
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(target), target, "The selected editable target is not supported.");
        }
    }
}
