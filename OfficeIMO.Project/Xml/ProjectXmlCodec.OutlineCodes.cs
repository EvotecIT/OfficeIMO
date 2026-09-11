using System.Xml.Linq;

namespace OfficeIMO.Project;

internal static partial class ProjectXmlCodec {
    private static readonly string[] OutlineOrder = "Guid FieldID FieldName Alias Enterprise ShowIndent ResourceSubstitutionEnabled LeafOnly AllLevelsRequired OnlyTableValuesAllowed Masks Values".Split(' ');
    private static readonly string[] OutlineMaskOrder = "Level Type Length Separator".Split(' ');
    private static readonly string[] OutlineValueOrder = "ValueID FieldGUID ParentValueID Type IsCollapsed Value Description".Split(' ');
    private static readonly string[] OutlineSelectionOrder = "FieldID ValueID ValueGUID".Split(' ');

    private static void ReadOutlineCodes(ProjectDocument document, XElement root, ProjectLoadOptions options, ref int entities, CancellationToken token) {
        foreach (var node in Children(root, "OutlineCodes", "OutlineCode")) {
            token.ThrowIfCancellationRequested(); CheckEntities(++entities, options);
            var definition = document.OutlineCodes.Add(); Attach(document, definition, node);
            ReadFields(definition, node, document, ProjectXmlFields.OutlineCodeDefinition);
            foreach (var child in Children(node, "Masks", "Mask")) {
                token.ThrowIfCancellationRequested(); CheckEntities(++entities, options);
                var item = definition.Masks.Add(); Attach(document, item, child);
                ReadFields(item, child, document, ProjectXmlFields.OutlineCodeMask);
            }
            foreach (var child in Children(node, "Values", "Value")) {
                token.ThrowIfCancellationRequested(); CheckEntities(++entities, options);
                var item = definition.Values.Add(); Attach(document, item, child);
                ReadFields(item, child, document, ProjectXmlFields.OutlineCodeLookupValue);
            }
        }
    }
    private static void ReadOutlineSelections(ProjectCollection<ProjectCustomFieldValue> values, XElement node, ProjectLoadOptions options, ref int entities, CancellationToken token) {
        foreach (var child in node.Elements(node.Name.Namespace + "OutlineCode")) {
            token.ThrowIfCancellationRequested(); CheckEntities(++entities, options);
            var value = values.Add(); Attach(values.Document, value, child);
            ReadFields(value, child, values.Document, ProjectXmlFields.CustomFieldValue);
        }
    }
    private static void WriteOutlineCodes(ProjectDocument document, XElement root, CancellationToken token) {
        ReplaceContainer(root, "OutlineCodes", "OutlineCode", document.OutlineCodes.Select(definition => {
            token.ThrowIfCancellationRequested();
            var node = NewNode(document, definition, "OutlineCode");
            ProjectXmlFields.Write(definition, node, document, ProjectXmlFields.OutlineCodeDefinition, OutlineOrder);
            ReplaceContainer(node, "Masks", "Mask", definition.Masks.Select(mask => {
                token.ThrowIfCancellationRequested(); var child = NewNode(document, mask, "Mask");
                ProjectXmlFields.Write(mask, child, document, ProjectXmlFields.OutlineCodeMask, OutlineMaskOrder); return child;
            }), OutlineOrder);
            ReplaceContainer(node, "Values", "Value", definition.Values.Select(value => {
                token.ThrowIfCancellationRequested(); var child = NewNode(document, value, "Value");
                ProjectXmlFields.Write(value, child, document, ProjectXmlFields.OutlineCodeLookupValue, OutlineValueOrder); return child;
            }), OutlineOrder);
            return node;
        }), RootOrder);
    }
    private static void WriteOutlineSelections(ProjectCollection<ProjectCustomFieldValue> values, XElement node, string[] order, CancellationToken token) {
        ReplaceChildren(node, "OutlineCode", values.Select(value => {
            token.ThrowIfCancellationRequested(); var child = NewNode(values.Document, value, "OutlineCode");
            ProjectXmlFields.Write(value, child, values.Document, ProjectXmlFields.CustomFieldValue, OutlineSelectionOrder); return child;
        }), order);
    }
}
