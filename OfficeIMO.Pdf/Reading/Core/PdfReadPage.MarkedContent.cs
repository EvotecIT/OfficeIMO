using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    /// <summary>
    /// Finds uniquely identified page scopes whose complete content cannot produce text.
    /// An absent run alone cannot distinguish an empty cell from a broken reference.
    /// </summary>
    internal Dictionary<int, int?> GetTextlessPageMarkedContent(Action consumeWork, CancellationToken cancellationToken) {
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        PageContentStreamSequence sequence = GetContentStreamSequence(new PageContentBudget(this, cancellationToken));
        var open = new Stack<(int? Id, int? Stream, long TextVersion)>();
        var seen = new HashSet<int>();
        var duplicates = new HashSet<int>();
        var textless = new Dictionary<int, int?>();
        long textVersion = 0;
        PdfContentStreamInterpreter.Interpret(sequence.Content, _limits.MaxContentOperations, operation => {
            consumeWork();
            if (operation.HasInvalidOperands ||
                operation.Name is "Tj" or "TJ" or "'" or "\"" or "Do" or "BI" ||
                !PdfContentOperators.IsStandard(operation.Name)) {
                textVersion++;
            }

            switch (operation.Name) {
                case "BDC": {
                    long startVersion = textVersion;
                    object? properties = !operation.HasInvalidOperands && operation.Operands.Count == 2
                        ? operation.Operands[1]
                        : null;
                    int? id = null;
                    if (properties is PdfContentDictionary inline) {
                        if (inline.Items.ContainsKey("ActualText")) textVersion++;
                        if (inline.Items.TryGetValue("MCID", out object? value) &&
                            value is double number && number >= 0D && number <= int.MaxValue && Math.Truncate(number) == number) {
                            id = (int)number;
                        }
                    } else if (properties is string name) {
                        if (GetMarkedContentActualTextBytes(resources, name) is not null) textVersion++;
                        id = GetMarkedContentMcid(resources, name);
                    }
                    if (id is int mcid && !seen.Add(mcid)) duplicates.Add(mcid);
                    open.Push((id, sequence.GetObjectNumber(operation.OperatorOffset), startVersion));
                    break;
                }
                case "BMC":
                    open.Push((null, null, textVersion));
                    break;
                case "EMC": {
                    if (open.Count == 0) break;
                    (int? id, int? stream, long startVersion) = open.Pop();
                    if (id.HasValue && startVersion == textVersion && !duplicates.Contains(id.Value)) {
                        textless.Add(id.Value, stream);
                    }
                    break;
                }
            }
        },
            inlineImageComponentCount: name => GetDeclaredColorSpaceComponentCount(resources, name),
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands,
            dispatchInvalidOperations: true,
            operandBuffer: new List<object>(8));
        foreach (int duplicate in duplicates) textless.Remove(duplicate);
        return textless;
    }
}
