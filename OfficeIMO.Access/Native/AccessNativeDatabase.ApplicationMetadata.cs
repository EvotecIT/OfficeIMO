using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Access;

internal sealed partial class AccessNativeDatabase {
    private void LoadDataMacros(CancellationToken cancellation) {
        var macros = new List<AccessDataMacroInfo>();
        foreach (var entry in _document.Catalog.Where(x => x.NativeType == 1 && x.NativePayloads.ContainsKey("LvExtra"))) {
            var metadata = new AccessTable(_document, entry.Name);
            LoadProperties(metadata, entry.NativePayloads["LvExtra"].GetBytes(), cancellation, macroMap: true);
            foreach (var property in metadata.Properties) {
                if (!(property.Value is string xml)) continue;
                try {
                    using var source = new StringReader(xml);
                    using var reader = XmlReader.Create(source, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = MaxMetadataBytes });
                    using var bounded = new OfficeXmlLimitingReader(reader, "Access data macro", 64, MaxCatalogObjects, MaxCatalogObjects, cancellation);
                    var root = XElement.Load(bounded, LoadOptions.PreserveWhitespace);
                    if (root.Name.NamespaceName != "http://schemas.microsoft.com/office/accessservices/2009/11/application" || root.Name.LocalName != "DataMacro") continue;
                    macros.Add(new AccessDataMacroInfo(entry, property.Key, xml, root.Descendants().Where(x => x.Name.LocalName == "Action" || x.Name.LocalName == "Comment").Select(x => x.Name.LocalName == "Action" ? (string?)x.Attribute("Name") ?? "Action" : "Comment").ToArray()));
                } catch (XmlException) { AddApplicationDiagnostic("access.data-macro.opaque", "An unqualified data-macro XML representation remains in exact native catalog payloads."); }
            }
        }
        _document.DataMacros = Array.AsReadOnly(macros.ToArray());
    }
    private void LoadResources(CancellationToken cancellation) {
        if (!_tables.TryGetValue("MSysResources", out var table)) return;
        RequireFields(table, "Id", "Name", "Type", "Extension", "Data");
        var resources = new List<AccessResourceInfo>();
        using var rows = new AccessNativeRowCursor(table, cancellation, rowLimit: MaxCatalogObjects);
        while (rows.Read(cancellation)) resources.Add(new AccessResourceInfo(Convert.ToInt32(RequiredField(table, rows, "Id", cancellation)),
            Field(table, rows, "Name", cancellation) as string, Field(table, rows, "Type", cancellation) as string,
            Field(table, rows, "Extension", cancellation) as string, Field(table, rows, "Data", cancellation) as AccessComplexValue, rows.Current.NativeBytes()));
        _document.Resources = Array.AsReadOnly(resources.ToArray());
    }
}
