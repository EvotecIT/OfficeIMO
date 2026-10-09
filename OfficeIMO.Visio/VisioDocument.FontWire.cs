using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    // All model edits and native error collection finish before this copy-only wire projection.
    private void WriteNativeFontNames(XDocument documentXml, Package package,
        IEnumerable<(VisioPage Page, PackagePart Part, PackageRelationship Relationship)> pages,
        IReadOnlyList<PackageMasterEntry> masters, PackagePart pagesPart, PackagePart? mastersPart) {
        XElement root = documentXml.Root!;
        var entries = VisioNativeCellMetadata.Read(root.Elements());
        var outputs = new List<(PackagePart Part, string Scope)>();
        VisioFontWireCodec.DeclareLiteralNames(root, new[] { root });
        foreach (var page in pages) {
            string scope = VisioNativeCellMetadata.PageScope(page.Page.Id);
            RebindPart(page.Part, scope, xml => {
                var elements = ShapeElementsById(xml);
                var ids = AssignPageElementIdentifiers(page.Page);
                foreach (VisioShape shape in page.Page.AllShapes())
                    if (shape.NativeFontScope != null && elements.TryGetValue(VisioShapeFormulaReferences.NormalizeSheetId(ids[shape.Id]), out XElement? emitted) && emitted != null)
                        RebindShapeFontOutput(shape, emitted, root, scope, entries);
            });
        }
        foreach (PackageMasterEntry master in masters) {
            string scope = VisioNativeCellMetadata.MasterScope(master.Master.NameU);
            RebindPart(package.GetPart(new System.Uri("/visio/masters/master" + master.PartNumber + ".xml", System.UriKind.Relative)), scope,
                xml => { if (xml.Root != null) RebindMasterFontOutput(master.Master, xml.Root, root, scope, entries); });
        }
        RebindPart(pagesPart, "Pages");
        if (mastersPart != null) RebindPart(mastersPart, "Masters", xml => {
            foreach (PackageMasterEntry master in masters) {
                XElement? entry = xml.Root?.Elements(XName.Get("Master", VisioNamespace)).SingleOrDefault(element => (string?)element.Attribute("ID") == master.PackageId);
                if (entry != null) RebindMasterFontOutput(master.Master, entry, root, VisioNativeCellMetadata.MasterScope(master.Master.NameU), entries);
            }
        });
        VisioFontWireCodec codec = VisioFontWireCodec.WriteDocument(root);
        codec.EncodeCells(root, root, entries);
        foreach (var output in outputs) {
            XDocument xml = LoadPackageXml(output.Part, "Written Visio font names");
            if (xml.Root == null) continue;
            codec.EncodeCells(xml.Root, root, entries, output.Scope);
            Write(output.Part, xml);
        }

        // Retain one private XML part at a time. Font declaration needs a first pass,
        // but does not require keeping the complete drawing XML graph in memory.
        void RebindPart(PackagePart part, string scope, System.Action<XDocument>? rebind = null) {
            XDocument xml = LoadPackageXml(part, "Written Visio font ownership");
            if (xml.Root == null) return;
            rebind?.Invoke(xml);
            VisioFontWireCodec.DeclareLiteralNames(root, new[] { xml.Root });
            Write(part, xml); outputs.Add((part, scope));
        }
        static void Write(PackagePart part, XDocument xml) {
            using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(xml.Declaration + System.Environment.NewLine + xml.ToString(SaveOptions.DisableFormatting));
        }
    }
}
