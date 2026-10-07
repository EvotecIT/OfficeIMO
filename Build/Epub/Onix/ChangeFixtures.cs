using OfficeIMO.Workflows;
using System.Xml.Linq;
using System.Xml.Schema;

internal static class ChangeFixtures {
    internal static void Write(string profile, BookOnixExportResult source, XmlSchemaSet schemas, string output) {
        XNamespace ns = BookProject.OnixNamespace;
        if (profile == "collateral-xhtml") {
            var message = BookOnixMessage.CreateBlockUpdates([new(source) { ReplaceBlocks = [BookOnixBlock.CollateralDetail] }], schemas);
            var before = XDocument.Load(new MemoryStream(source.Bytes)).Descendants(ns + "CollateralDetail").Single();
            var after = XDocument.Load(new MemoryStream(message.Bytes)).Descendants(ns + "CollateralDetail").Single();
            if (!XNode.DeepEquals(before, after)) throw new InvalidDataException("A replacement lost collateral content or whitespace.");
            Save("update-collateral", message);
        }
        if (profile == "advance") {
            var message = BookOnixMessage.CreateBlockUpdates([new(source) { ReplaceBlocks = [BookOnixBlock.ProductSupply] }], schemas);
            var before = XDocument.Load(new MemoryStream(source.Bytes)).Descendants(ns + "ProductSupply").ToArray();
            var after = XDocument.Load(new MemoryStream(message.Bytes)).Descendants(ns + "ProductSupply").ToArray();
            if (before.Length != 2 || after.Length != 2 || before.Where((block, index) => !XNode.DeepEquals(block, after[index])).Any())
                throw new InvalidDataException("A supply update did not preserve both complete markets.");
            Save("update-supply", message);
        }
        if (profile == "confirmed") {
            Save("update-clears", BookOnixMessage.CreateBlockUpdates([new(source) { ClearBlocks = [
                BookOnixBlock.CollateralDetail, BookOnixBlock.PromotionDetail, BookOnixBlock.ContentDetail,
                BookOnixBlock.RelatedMaterial, BookOnixBlock.ProductionDetail] }], schemas));
            Save("delete-record", BookOnixMessage.CreateDeletions([new(source) { Reasons = [
                new("Record issued in error & duplicated", "eng"), new("Błędny rekord", "pol")] }], schemas));
            Save("delete-record-no-reason", BookOnixMessage.CreateDeletions([new(source)], schemas));
        }
        void Save(string name, BookOnixMessage message) => File.WriteAllBytes(Path.Combine(output, name + ".onix"), message.Bytes);
    }

    internal static void WriteCatalog(IReadOnlyList<BookOnixExportResult> sources, XmlSchemaSet schemas, string output) {
        File.WriteAllBytes(Path.Combine(output, "update-catalog.onix"), BookOnixMessage.CreateBlockUpdates(sources.Select(source =>
            new BookOnixBlockUpdate(source) { ReplaceBlocks = [BookOnixBlock.DescriptiveDetail, BookOnixBlock.PublishingDetail] }).ToArray(), schemas).Bytes);
        File.WriteAllBytes(Path.Combine(output, "delete-catalog.onix"), BookOnixMessage.CreateDeletions(sources.Select(source =>
            new BookOnixDeletion(source)).ToArray(), schemas).Bytes);
    }
}
