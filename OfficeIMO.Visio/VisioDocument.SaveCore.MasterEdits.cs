using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private XElement CreateMasterModelShapeXml(VisioShape shape, XDocument? source = null, IReadOnlyDictionary<string, int>? layerIndexes = null) {
        var ids = BuildPersistedIdMap(new[] { shape }, Array.Empty<VisioConnector>(), new Dictionary<string, VisioMaster>(),
            source?.Descendants(XName.Get("Shape", VisioNamespace)).Attributes("ID").Select(attribute => attribute.Value));
        var document = new XDocument();
        using (var writer = document.CreateWriter()) {
            WriteShapeElement(writer, VisioNamespace, shape, ids,
                new Dictionary<string, VisioMaster>(), Array.Empty<PackageMasterEntry>(), layerIndexes ?? new Dictionary<string, int>());
        }
        // A change snapshot must retain actual model values. Authoring defaults in the ordinary
        // writer (zero width -> 1, zero local pin -> center) must not conceal a later explicit edit.
        void CaptureTransform(VisioShape model, XElement xml) {
            foreach (var value in new[] { ("Width", model.Width), ("Height", model.Height), ("LocPinX", model.LocPinX), ("LocPinY", model.LocPinY) }) {
                XElement? cell = xml.Elements(XName.Get("Cell", VisioNamespace)).FirstOrDefault(e => (string?)e.Attribute("N") == value.Item1);
                if (cell != null) cell.SetAttributeValue("V", XmlConvert.ToString(value.Item2));
            }
            var children = xml.Element(XName.Get("Shapes", VisioNamespace))?.Elements(XName.Get("Shape", VisioNamespace)).ToArray();
            if (children != null) for (int i = 0; i < model.Children.Count; i++) CaptureTransform(model.Children[i], children[i]);
        }
        CaptureTransform(shape, document.Root!);
        return document.Root!;
    }

    private void ApplyLoadedMasterShapeChanges(XDocument content, VisioMaster master) {
        if (master.LoadedModelShapeXml == null) return;
        XElement? source = content.Root?.Element(XName.Get("Shapes", VisioNamespace))?.Elements(XName.Get("Shape", VisioNamespace)).FirstOrDefault();
        if (source == null) return;
        MergeModeledContentChanges(source, master.LoadedModelShapeXml, CreateMasterModelShapeXml(master.Shape, master.RawMasterContentXml));
    }

}
