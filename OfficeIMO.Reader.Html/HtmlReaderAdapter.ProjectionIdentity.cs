using System.IO;

namespace OfficeIMO.Reader.Html;

internal static partial class HtmlReaderAdapter {
    internal static void PrefixProjection(string prefix, OfficeDocumentReadResult result, ref int tableIndex) {
        foreach (OfficeDocumentBlock block in result.Blocks) {
            block.Id = prefix + block.Id;
            if (!string.IsNullOrWhiteSpace(block.Location.BlockAnchor)) block.Location.BlockAnchor = prefix + block.Location.BlockAnchor;
        }
        foreach (ReaderTable table in result.Tables) {
            if (table.Location != null) {
                table.Location.TableIndex = tableIndex++;
                if (!string.IsNullOrWhiteSpace(table.Location.BlockAnchor)) table.Location.BlockAnchor = prefix + table.Location.BlockAnchor;
            }
        }
        foreach (OfficeDocumentLink link in result.Links) {
            link.Id = prefix + link.Id;
            if (!string.IsNullOrWhiteSpace(link.Location.BlockAnchor)) link.Location.BlockAnchor = prefix + link.Location.BlockAnchor;
        }
        foreach (OfficeDocumentFormField form in result.Forms) {
            form.Id = prefix + form.Id;
            if (!string.IsNullOrWhiteSpace(form.Location.BlockAnchor)) form.Location.BlockAnchor = prefix + form.Location.BlockAnchor;
        }
        foreach (OfficeDocumentAsset asset in result.Assets) {
            asset.Id = prefix + asset.Id;
            string? extension = string.IsNullOrWhiteSpace(asset.Extension)
                ? Path.GetExtension(asset.FileName)
                : asset.Extension;
            asset.FileName = OfficeDocumentAssetNaming.BuildFileName(asset.Id, extension);
            if (!string.IsNullOrWhiteSpace(asset.Location.BlockAnchor)) asset.Location.BlockAnchor = prefix + asset.Location.BlockAnchor;
        }
        foreach (ReaderVisual visual in result.Visuals) {
            if (visual.Location != null && !string.IsNullOrWhiteSpace(visual.Location.BlockAnchor)) visual.Location.BlockAnchor = prefix + visual.Location.BlockAnchor;
        }
    }
}
