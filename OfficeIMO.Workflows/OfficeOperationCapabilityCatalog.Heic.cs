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
                "Inserted as PNG using the registered platform codec. Requires ImageIO on Apple or WIC with the HEIF and HEVC extensions on Windows; no bundled decoder or Linux codec. HDR gain maps and depth are not retained."));
        }
    }
}
