using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static void ReconcileMergeHeadStyles(XElement first, XElement second, EpubChapterMergeStylePolicy policy, CancellationToken token) {
        XElement left = ComparableMergeHead(first), right = ComparableMergeHead(second);
        if (policy == EpubChapterMergeStylePolicy.RequireEquivalent) {
            if (!XNode.DeepEquals(left, right))
                throw new NotSupportedException("Resolve conflicting chapter heads or select an explicit stylesheet reconciliation policy.");
            return;
        }
        // Compare everything except the opted-in stylesheet nodes. Head attributes,
        // metadata, comments and processing instructions are still preservation boundaries.
        left.Elements().Where(IsMergeHeadStyle).Remove();
        right.Elements().Where(IsMergeHeadStyle).Remove();
        if (!XNode.DeepEquals(left, right))
            throw new NotSupportedException("Appending styles requires otherwise equivalent chapter heads.");
        foreach (XElement style in second.Elements().Where(IsMergeHeadStyle)) {
            token.ThrowIfCancellationRequested();
            // Do not deduplicate: repeated stylesheet declarations and their position can affect the cascade.
            first.Add(new XElement(style));
        }
    }

    private static bool IsMergeHeadStyle(XElement element) => element.Name == Html + "style" ||
        element.Name == Html + "link" && Tokens((string?)element.Attribute("rel"))
            .Any(value => string.Equals(value, "stylesheet", StringComparison.OrdinalIgnoreCase));
}
