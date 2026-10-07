using OfficeIMO.Epub;
using OfficeIMO.Workflows;
using System.Xml.Linq;
using System.Xml.Schema;

internal static class MarketFixtures {
    internal static void Write(BookProject project, BookOnixExportOptions options, XmlSchemaSet schemas, string output, DateTimeOffset timestamp) {
        var supplies = options.Commercial!.Supplies.Select((supply, index) => supply with { MarketReference = "market-" + index }).ToArray();
        var source = project.ExportOnix(options with { Commercial = options.Commercial with { Supplies = supplies } }, schemas,
            new EpubWriteOptions { ModifiedAt = timestamp });
        File.WriteAllBytes(Path.Combine(output, "named-markets.onix"), source.Bytes);
        File.WriteAllBytes(Path.Combine(output, "named-markets.epub"), source.Publication.Bytes);
        var update = BookOnixMessage.CreateBlockUpdates([new(source) { ReplaceMarketReferences = ["market-1"], RemoveMarketReferences = ["market-0"] }], schemas);
        XNamespace ns = BookProject.OnixNamespace;
        var blocks = XDocument.Load(new MemoryStream(update.Bytes)).Descendants(ns + "ProductSupply").ToArray();
        var original = XDocument.Load(new MemoryStream(source.Bytes)).Descendants(ns + "ProductSupply").Last();
        if (blocks.Length != 2 || !XNode.DeepEquals(original, blocks[0]) || blocks[1].Elements().Count() != 1 ||
            blocks[1].Element(ns + "MarketReference")?.Value != "market-0") throw new InvalidDataException("Market update/removal contract failed.");
        File.WriteAllBytes(Path.Combine(output, "update-named-markets.onix"), update.Bytes);
        File.WriteAllBytes(Path.Combine(output, "remove-absent-market.onix"), BookOnixMessage.CreateBlockUpdates([
            new(source) { RemoveMarketReferences = ["former-market"] }], schemas).Bytes);
    }
}
