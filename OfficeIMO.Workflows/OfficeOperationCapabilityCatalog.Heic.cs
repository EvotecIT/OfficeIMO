namespace OfficeIMO.Workflows;

public static partial class OfficeOperationCapabilityCatalog {
    private static void AddHeicInsertionRows(ICollection<OfficeOperationCapability> rows) {
        Add("OfficeIMO.Word", "DOCX", "WordParagraph.AddImage");
        Add("OfficeIMO.PowerPoint", "PPTX", "PowerPointSlide.AddPicture");
        Add("OfficeIMO.Excel", "XLSX", "ExcelSheet.AddImage");

        void Add(string packageId, string target, string api) {
            rows.Add(new OfficeOperationCapability(
                "image:heic:" + target.ToLowerInvariant() + ":insert",
                packageId, "HEIC", "heic-insert", OfficeOperationKind.Convert,
                OfficeOperationSupportState.Partial, api,
                "OfficeImageReader.NormalizeForEmbedding", "OfficeIMO.Operations",
                new[] { ".heic", ".heif", ".hif" }, target,
                "Inserted as PNG through an explicitly registered .NET raster codec. Use OfficeIMO.Drawing.ImageIO on macOS/iOS or supply your own codec elsewhere. No codec is registered automatically and no HEVC decoder is bundled. HDR gain maps and depth are not retained."));
        }
    }
}
