using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

namespace OfficeIMO.Workflows;

public sealed partial class BookOnixMessage {
    /// <summary>
    /// Composes 1–1000 explicit block updates (notification 04). Selected blocks are copied whole;
    /// unselected blocks are omitted, and only ClearBlocks emits empty blocks. All source headers must match.
    /// ProductSupply copies every source repeat, without per-market selection. Missing selected blocks fail.
    /// Retains the source EPUB association and the same integrity, schema, cancellation and size gates as Create.
    /// Does not compare recipient state or transmit instructions. Do not mutate inputs or schemas during composition.
    /// </summary>
    public static BookOnixMessage CreateBlockUpdates(IReadOnlyList<BookOnixBlockUpdate> updates, XmlSchemaSet schemas,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(updates);
        cancellationToken.ThrowIfCancellationRequested();
        if (updates.Count is < 1 or > 1000) throw new ArgumentException("Supply 1–1000 block updates.", nameof(updates));
        var sources = new BookOnixExportResult[updates.Count];
        var replacements = new HashSet<BookOnixBlock>[updates.Count];
        var clears = new HashSet<BookOnixBlock>[updates.Count];
        for (int index = 0; index < updates.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            var update = updates[index]; ArgumentNullException.ThrowIfNull(update); ArgumentNullException.ThrowIfNull(update.Source);
            sources[index] = update.Source;
            replacements[index] = ReadBlocks(update.ReplaceBlocks);
            clears[index] = ReadBlocks(update.ClearBlocks);
            if (replacements[index].Count + clears[index].Count == 0 || replacements[index].Overlaps(clears[index]))
                throw new ArgumentException("Select at least one block; replace and clear cannot overlap.", nameof(updates));
            if (clears[index].Any(block => block is BookOnixBlock.DescriptiveDetail or BookOnixBlock.PublishingDetail or BookOnixBlock.ProductSupply))
                throw new ArgumentException("Only optional resettable blocks may be cleared.", nameof(updates));
        }
        return Compose(sources, schemas, BookOnixMessageKind.BlockUpdates, (index, record) => {
            XNamespace ns = BookProject.OnixNamespace;
            var result = CopyRecordIdentity(record, "04");
            foreach (var block in Enum.GetValues<BookOnixBlock>()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (clears[index].Contains(block)) result.Add(new XElement(ns + block.ToString()));
                else if (replacements[index].Contains(block)) {
                    var elements = record.Elements(ns + block.ToString()).ToArray();
                    if (elements.Length == 0) throw new ArgumentException("Selected replacement block is absent: " + block + ". Use an explicit clear where permitted.", nameof(updates));
                    // The current exporter does not author MarketReference. Guard this boundary if it is extended later.
                    if (block == BookOnixBlock.ProductSupply && elements.Any(element => element.Element(ns + "MarketReference") != null))
                        throw new ArgumentException("Market-referenced supply needs an explicit per-market update contract.", nameof(updates));
                    result.Add(elements.Select(element => new XElement(element)));
                }
            }
            return result;
        }, cancellationToken);
    }

    /// <summary>
    /// Composes explicit metadata-record deletions (notification 05), keeping record references and product identifiers.
    /// Reasons are optional plain text; product status changes must instead use complete records or block updates.
    /// The source headers, exact EPUB provenance, integrity and size gates remain in force. No deletion is transmitted.
    /// </summary>
    public static BookOnixMessage CreateDeletions(IReadOnlyList<BookOnixDeletion> deletions, XmlSchemaSet schemas,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(deletions);
        cancellationToken.ThrowIfCancellationRequested();
        if (deletions.Count is < 1 or > 1000) throw new ArgumentException("Supply 1–1000 deletions.", nameof(deletions));
        XNamespace ns = BookProject.OnixNamespace;
        var sources = new BookOnixExportResult[deletions.Count];
        var reasons = new XElement[deletions.Count][];
        for (int index = 0; index < deletions.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            var deletion = deletions[index]; ArgumentNullException.ThrowIfNull(deletion); ArgumentNullException.ThrowIfNull(deletion.Source);
            sources[index] = deletion.Source;
            ArgumentNullException.ThrowIfNull(deletion.Reasons);
            if (deletion.Reasons.Count > 16) throw new ArgumentException("At most 16 deletion-reason translations are supported.", nameof(deletions));
            var languages = new HashSet<string>(StringComparer.Ordinal); var values = new List<XElement>();
            foreach (var reason in deletion.Reasons) {
                cancellationToken.ThrowIfCancellationRequested();
                ArgumentNullException.ThrowIfNull(reason); ArgumentException.ThrowIfNullOrWhiteSpace(reason.Text);
                if (reason.Text.Length > 100) throw new ArgumentException("Deletion reasons cannot exceed 100 UTF-16 code units.", nameof(deletions));
                XmlConvert.VerifyXmlChars(reason.Text);
                BookProject.RequireOnixTranslationLanguage(reason.LanguageCode, deletion.Reasons.Count, nameof(deletion.Reasons));
                if (!languages.Add(reason.LanguageCode ?? string.Empty)) throw new ArgumentException("Deletion reason languages must be distinct.", nameof(deletions));
                values.Add(new XElement(ns + "DeletionText", reason.LanguageCode != null ? new XAttribute("language", reason.LanguageCode) : null, reason.Text));
            }
            reasons[index] = values.ToArray();
        }
        return Compose(sources, schemas, BookOnixMessageKind.Deletions, (index, record) => {
            var result = CopyRecordIdentity(record, "05");
            result.Element(ns + "NotificationType")!.AddAfterSelf(reasons[index]);
            return result;
        }, cancellationToken);
    }

    private static HashSet<BookOnixBlock> ReadBlocks(IReadOnlyList<BookOnixBlock> blocks) {
        ArgumentNullException.ThrowIfNull(blocks);
        if (blocks.Count > 8) throw new ArgumentException("At most eight blocks may be selected.", nameof(blocks));
        var result = new HashSet<BookOnixBlock>();
        foreach (var block in blocks)
            if (!Enum.IsDefined(block) || !result.Add(block)) throw new ArgumentException("Block selections must be defined and distinct.", nameof(blocks));
        return result;
    }

    private static XElement CopyRecordIdentity(XElement record, string notification) {
        XNamespace ns = BookProject.OnixNamespace;
        // ExportOnix owns this complete-record shape. Copy block zero in full instead of reconstructing an ISBN-only identity.
        var names = Enum.GetNames<BookOnixBlock>().ToHashSet(StringComparer.Ordinal);
        var result = new XElement(record.Name, record.Attributes(), record.Elements().TakeWhile(element => !names.Contains(element.Name.LocalName))
            .Where(element => element.Name != ns + "DeletionText").Select(element => new XElement(element)));
        result.Element(ns + "NotificationType")!.Value = notification;
        return result;
    }
}
