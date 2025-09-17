using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Collections.Generic;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a SmartArt diagram in a <see cref="WordDocument"/>.
    /// </summary>
    public class WordSmartArt : WordElement {
        private static int _docPrIdSeed = 1;

        private static UInt32Value GenerateDocPrId() {
            int id = Interlocked.Increment(ref _docPrIdSeed);
            return (UInt32Value)(uint)id;
        }

        internal Drawing _drawing = null!;
        private readonly WordDocument _document;
        private readonly WordParagraph _paragraph;
        private readonly SmartArtType? _type;

        internal WordSmartArt(WordDocument document, WordParagraph paragraph, SmartArtType type) {
            _document = document;
            _paragraph = paragraph;
            _type = type;

            InsertSmartArt(type);
        }

        internal WordSmartArt(WordDocument document, WordParagraph paragraph, Drawing drawing) {
            _document = document;
            _paragraph = paragraph;
            _drawing = drawing;
            _type = TryDetectType();
        }

        private void InsertSmartArt(SmartArtType type) {
            var mainPart = _document._wordprocessingDocument.MainDocumentPart!;

            // Build SmartArt parts in-memory (no external templates)
            var (relLayout, relColors, relStyle, relData) = SmartArtBuiltIn.AddParts(mainPart, type);

            var graphic = new Graphic(new GraphicData(
                new RelationshipIds {
                    LayoutPart = relLayout,
                    StylePart = relStyle,
                    ColorPart = relColors,
                    DataPart = relData
                }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram" });

            // Match exported templates effect extents for visual parity
            EffectExtent eff;
            if (type == SmartArtType.CustomSmartArt1) {
                eff = new EffectExtent { LeftEdge = 38100L, TopEdge = 0L, RightEdge = 57150L, BottomEdge = 0L };
            } else if (type == SmartArtType.CustomSmartArt2) {
                eff = new EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 19050L };
            } else {
                eff = new EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            }

            var inline = new Inline(
                new Extent { Cx = 5486400, Cy = 3200400 },
                eff,
                new DocProperties { Id = GenerateDocPrId(), Name = "Diagram 1" },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                    new GraphicFrameLocks { NoChangeAspect = true }),
                graphic) {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                };

            _drawing = new Drawing(inline);
            _paragraph.VerifyRun();
            _paragraph._run.Append(_drawing);
        }

        // All SmartArt parts are created in SmartArtBuiltIn.

        /// <summary>
        /// Gets the number of editable node paragraphs in the SmartArt.
        /// Nodes are detected as points that have a text body and no explicit type (not doc/parTrans/sibTrans/pres).
        /// </summary>
        public int NodeCount {
            get {
                var (xdoc, _, paras) = LoadNodeParagraphs();
                return paras.Count;
            }
        }

        /// <summary>
        /// Returns the text of the node at the given index (0-based).
        /// </summary>
        public string GetNodeText(int index) {
            var (xdoc, ns, paras) = LoadNodeParagraphs();
            if (index < 0 || index >= paras.Count) throw new ArgumentOutOfRangeException(nameof(index));
            var a = ns.a;
            var p = paras[index];
            var text = string.Concat(p.Elements(a + "r").Elements(a + "t").Select(t => (string)t));
            return text ?? string.Empty;
        }

        /// <summary>
        /// Sets the text of the node at the given index (0-based). Preserves end paragraph run properties.
        /// </summary>
        public void SetNodeText(int index, string text) {
            var (xdoc, ns, paras, dataPart) = LoadNodeParagraphsWithPart();
            if (index < 0 || index >= paras.Count) throw new ArgumentOutOfRangeException(nameof(index));
            ReplaceParagraphText(paras[index], ns.a, text);
            SaveDiagramData(dataPart, xdoc);
        }

        /// <summary>
        /// Replaces texts of all nodes in order. If more texts provided than nodes, extras are ignored.
        /// If fewer texts provided, remaining nodes are left unchanged.
        /// </summary>
        public void ReplaceTexts(IEnumerable<string> texts) {
            var (xdoc, ns, paras, dataPart) = LoadNodeParagraphsWithPart();
            int i = 0;
            foreach (var t in texts) {
                if (i >= paras.Count) break;
                ReplaceParagraphText(paras[i], ns.a, t);
                i++;
            }
            SaveDiagramData(dataPart, xdoc);
        }

        /// <summary>
        /// Convenience overload for replacing texts.
        /// </summary>
        public void ReplaceTexts(params string[] texts) => ReplaceTexts((IEnumerable<string>)texts);

        /// <summary>
        /// Replaces texts of all nodes with optional formatting applied uniformly to each replacement.
        /// If more texts are provided than nodes, extras are ignored. If fewer, remaining nodes are unchanged.
        /// </summary>
        public void ReplaceTexts(IEnumerable<string> texts, bool bold, bool italic, bool underline, string? colorHex = null, double? sizePt = null) {
            var (xdoc, ns, paras, dataPart) = LoadNodeParagraphsWithPart();
            int i = 0;
            foreach (var t in texts) {
                if (i >= paras.Count) break;
                ReplaceParagraphText(paras[i], ns.a, t, bold, italic, underline, colorHex, sizePt);
                i++;
            }
            SaveDiagramData(dataPart, xdoc);
        }

        /// <summary>
        /// Convenience overload to apply formatting to each provided text.
        /// </summary>
        public void ReplaceTexts(bool bold, bool italic, bool underline, string? colorHex = null, double? sizePt = null, params string[] texts)
            => ReplaceTexts((IEnumerable<string>)texts, bold, italic, underline, colorHex, sizePt);

        /// <summary>
        /// Sets node text with optional basic formatting and newline support.
        /// </summary>
        public void SetNodeText(int index, string text, bool bold, bool italic) {
            var (xdoc, ns, paras, dataPart) = LoadNodeParagraphsWithPart();
            if (index < 0 || index >= paras.Count) throw new ArgumentOutOfRangeException(nameof(index));
            ReplaceParagraphText(paras[index], ns.a, text, bold, italic);
            SaveDiagramData(dataPart, xdoc);
        }

        /// <summary>
        /// Sets the text of the node with extended formatting.
        /// </summary>
        public void SetNodeText(int index, string text, bool bold, bool italic, bool underline, string? colorHex, double? sizePt) {
            var (xdoc, ns, paras, dataPart) = LoadNodeParagraphsWithPart();
            if (index < 0 || index >= paras.Count) throw new ArgumentOutOfRangeException(nameof(index));
            ReplaceParagraphText(paras[index], ns.a, text, bold, italic, underline, colorHex, sizePt);
            SaveDiagramData(dataPart, xdoc);
        }

        /// <summary>
        /// Adds a new node (algorithmic layouts only: BasicProcess or Cycle) and sets its text.
        /// </summary>
        public void AddNode(string text) {
            if (!CanModifyNodes) throw new NotSupportedException("Adding nodes is supported only for algorithmic layouts (BasicProcess, Cycle). Use formatted templates for exact visual parity with fixed nodes.");
            var (xdoc, ns, paras, dataPart) = LoadNodeParagraphsWithPart();
            var dgm = ns.dgm; var a = ns.a;

            var ptLst = xdoc.Descendants(dgm + "ptLst").FirstOrDefault();
            var cxnLst = xdoc.Descendants(dgm + "cxnLst").FirstOrDefault();
            if (ptLst == null) throw new InvalidOperationException("SmartArt data model missing ptLst.");
            if (cxnLst == null) { cxnLst = new XElement(dgm + "cxnLst"); xdoc.Root?.Add(cxnLst); }

            // Find document node id
            var docPt = xdoc.Descendants(dgm + "pt").FirstOrDefault(p => (string?)p.Attribute("type") == "doc");
            if (docPt == null) throw new InvalidOperationException("Cannot locate document point in SmartArt data model.");
            var docId = (string?)docPt.Attribute("modelId") ?? throw new InvalidOperationException("Document point missing modelId.");

            // Create node point
            string newId = "{" + Guid.NewGuid().ToString().ToUpper() + "}";
            var p = new XElement(a + "p", new XElement(a + "endParaRPr",
                new XAttribute("lang", "en-US")));
            // In data model, the text container element is 'dgm:t' (alias of txBody)
            var txBody = new XElement(dgm + "t",
                new XElement(a + "bodyPr"),
                new XElement(a + "lstStyle"),
                p);
            var pt = new XElement(dgm + "pt",
                new XAttribute("modelId", newId),
                new XAttribute("type", "node"),
                new XElement(dgm + "prSet",
                    new XAttribute("placeholder", 1),
                    new XAttribute("phldrT", "[Text]")),
                new XElement(dgm + "spPr"),
                txBody);
            ptLst.Add(pt);

            // Determine next position
            var existing = cxnLst.Elements(dgm + "cxn").Where(x => (string?)x.Attribute("srcId") == docId).ToList();
            uint nextPos = existing.Any() ? (uint)(existing.Select(x => (int?)x.Attribute("srcOrd") ?? 0).DefaultIfEmpty(0).Max() + 1) : 0U;

            // Create connection
            var cxn = new XElement(dgm + "cxn",
                new XAttribute("modelId", "{" + Guid.NewGuid().ToString().ToUpper() + "}"),
                new XAttribute("srcId", docId),
                new XAttribute("destId", newId),
                new XAttribute("srcOrd", nextPos),
                new XAttribute("destOrd", 0));
            cxnLst.Add(cxn);

            // Set text with formatting preserved default
            ReplaceParagraphText(p, a, text);
            SaveDiagramData(dataPart, xdoc);
        }

        /// <summary>
        /// Inserts a node at a specific index (algorithmic layouts only). Re-sequences following nodes.
        /// </summary>
        public void InsertNodeAt(int index, string text) {
            if (!CanModifyNodes) throw new NotSupportedException("Inserting nodes is supported only for algorithmic layouts (BasicProcess, Cycle).");
            var (xdoc, ns, paras, dataPart) = LoadNodeParagraphsWithPart();
            var dgm = ns.dgm; var a = ns.a;
            var ptLst = xdoc.Descendants(dgm + "ptLst").FirstOrDefault();
            var cxnLst = xdoc.Descendants(dgm + "cxnLst").FirstOrDefault();
            if (ptLst == null) throw new InvalidOperationException("SmartArt data model missing ptLst.");
            if (cxnLst == null) { cxnLst = new XElement(dgm + "cxnLst"); xdoc.Root?.Add(cxnLst); }
            var docPt = xdoc.Descendants(dgm + "pt").FirstOrDefault(p => (string?)p.Attribute("type") == "doc");
            if (docPt == null) throw new InvalidOperationException("Cannot locate document point in SmartArt data model.");
            var docId = (string?)docPt.Attribute("modelId") ?? throw new InvalidOperationException("Document point missing modelId.");

            // Resequence existing srcOrd >= index
            var conns = cxnLst.Elements(dgm + "cxn").Where(c => (string?)c.Attribute("srcId") == docId).OrderBy(c => (int?)c.Attribute("srcOrd") ?? 0).ToList();
            if (index < 0 || index > conns.Count) throw new ArgumentOutOfRangeException(nameof(index));
            foreach (var c in conns.Where((c, i) => i >= index)) {
                var cur = (int?)c.Attribute("srcOrd") ?? 0;
                c.SetAttributeValue("srcOrd", cur + 1);
            }

            // Create new point
            string newId = "{" + Guid.NewGuid().ToString().ToUpper() + "}";
            var p = new XElement(a + "p", new XElement(a + "endParaRPr",
                new XAttribute("lang", "en-US")));
            var txBody = new XElement(dgm + "t",
                new XElement(a + "bodyPr"),
                new XElement(a + "lstStyle"),
                p);
            var pt = new XElement(dgm + "pt",
                new XAttribute("modelId", newId),
                new XAttribute("type", "node"),
                new XElement(dgm + "prSet",
                    new XAttribute("placeholder", 1),
                    new XAttribute("phldrT", "[Text]")),
                new XElement(dgm + "spPr"),
                txBody);
            ptLst.Add(pt);

            // Create connection with srcOrd=index
            var cxn = new XElement(dgm + "cxn",
                new XAttribute("modelId", "{" + Guid.NewGuid().ToString().ToUpper() + "}"),
                new XAttribute("srcId", docId),
                new XAttribute("destId", newId),
                new XAttribute("srcOrd", (uint)index),
                new XAttribute("destOrd", 0));
            cxnLst.Add(cxn);

            ReplaceParagraphText(p, a, text);
            SaveDiagramData(dataPart, xdoc);
        }

        /// <summary>
        /// Removes all nodes (algorithmic layouts only). Diagram remains with no child nodes.
        /// </summary>
        public void ClearNodes() {
            if (!CanModifyNodes) throw new NotSupportedException("Clearing nodes is supported only for algorithmic layouts (BasicProcess, Cycle).");
            var (xdoc, ns, paras, dataPart) = LoadNodeParagraphsWithPart();
            var dgm = ns.dgm; var a = ns.a;
            var pts = xdoc.Descendants(dgm + "pt").ToList();
            // Consider both 't' and legacy 'txBody' spellings; include explicit type="node"
            var nodePts = pts.Where(p => {
                var typ = (string?)p.Attribute("type");
                return (typ == null || typ == "node") && (p.Element(dgm + "t") != null || p.Element(dgm + "txBody") != null);
            }).ToList();
            foreach (var pt in nodePts) pt.Remove();
            var docPt = xdoc.Descendants(dgm + "pt").FirstOrDefault(p => (string)p.Attribute("type") == "doc");
            var docId = (string?)docPt?.Attribute("modelId");
            var cxnLst = xdoc.Descendants(dgm + "cxnLst").FirstOrDefault();
            if (docId != null && cxnLst != null) {
                var toRemove = cxnLst.Elements(dgm + "cxn").Where(x => (string)x.Attribute("srcId") == docId).ToList();
                foreach (var c in toRemove) c.Remove();
            }
            SaveDiagramData(dataPart, xdoc);
        }

        /// <summary>
        /// Removes a node by index (algorithmic layouts only).
        /// </summary>
        public void RemoveNodeAt(int index) {
            if (!CanModifyNodes) throw new NotSupportedException("Removing nodes is supported only for algorithmic layouts (BasicProcess, Cycle).");
            var (xdoc, ns, paras, dataPart) = LoadNodeParagraphsWithPart();
            if (index < 0 || index >= paras.Count) throw new ArgumentOutOfRangeException(nameof(index));
            var dgm = ns.dgm; var a = ns.a;

            // Re-identify node pts to make sure we remove the correct one (same filter as LoadNodeParagraphs)
            var pts = xdoc.Descendants(dgm + "pt");
            var nodePts = pts.Where(p => { var typ = (string?)p.Attribute("type"); return (typ == null || typ == "node") && (p.Element(dgm + "t") != null || p.Element(dgm + "txBody") != null); }).ToList();
            var targetPt = nodePts[index];
            var targetId = (string)targetPt.Attribute("modelId")!;

            // Remove cxn entries
            var cxnLst = xdoc.Descendants(dgm + "cxnLst").FirstOrDefault();
            if (cxnLst != null) {
                var toRemove = cxnLst.Elements(dgm + "cxn").Where(x => (string)x.Attribute("destId") == targetId).ToList();
                foreach (var c in toRemove) c.Remove();
                // Resequence srcOrd for remaining doc->child connections
                var docPt = xdoc.Descendants(dgm + "pt").FirstOrDefault(p => (string)p.Attribute("type") == "doc");
                var docId = (string?)docPt?.Attribute("modelId");
                if (docId != null) {
                    var conns = cxnLst.Elements(dgm + "cxn").Where(x => (string)x.Attribute("srcId") == docId).ToList();
                    uint ord = 0;
                    foreach (var c in conns.OrderBy(c => (int?)c.Attribute("srcOrd") ?? 0)) {
                        c.SetAttributeValue("srcOrd", ord++);
                    }
                }
            }

            // Remove the point
            targetPt.Remove();
            SaveDiagramData(dataPart, xdoc);
        }

        public bool CanModifyNodes => _type == SmartArtType.BasicProcess || _type == SmartArtType.Cycle;

        private (XDocument xdoc, (XNamespace dgm, XNamespace a) ns, List<XElement> paras) LoadNodeParagraphs() {
            var dataPart = GetDiagramDataPart();
            var xdoc = LoadDiagramXDocument(dataPart);
            var dgm = (XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/diagram";
            var a = (XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/main";

            // Points with no @type and with a dgm:t (or dgm:txBody) are treated as editable nodes.
            var pts = xdoc.Descendants(dgm + "pt");
            var nodePts = pts.Where(p => {
                var typ = (string?)p.Attribute("type");
                return (typ == null || typ == "node") && (p.Element(dgm + "t") != null || p.Element(dgm + "txBody") != null);
            });
            var paras = nodePts
                .Select(p => (p.Element(dgm + "t") ?? p.Element(dgm + "txBody"))?.Element(a + "p"))
                .Where(p => p != null)
                .Cast<XElement>()
                .ToList();
            return (xdoc, (dgm, a), paras);
        }

        private (XDocument xdoc, (XNamespace dgm, XNamespace a) ns, List<XElement> paras, DiagramDataPart dataPart) LoadNodeParagraphsWithPart() {
            var dataPart = GetDiagramDataPart();
            var xdoc = LoadDiagramXDocument(dataPart);
            var dgm = (XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/diagram";
            var a = (XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/main";

            var pts = xdoc.Descendants(dgm + "pt");
            var nodePts = pts.Where(p => {
                var typ = (string?)p.Attribute("type");
                return (typ == null || typ == "node") && (p.Element(dgm + "t") != null || p.Element(dgm + "txBody") != null);
            });
            var paras = nodePts
                .Select(p => (p.Element(dgm + "t") ?? p.Element(dgm + "txBody"))?.Element(a + "p"))
                .Where(p => p != null)
                .Cast<XElement>()
                .ToList();
            return (xdoc, (dgm, a), paras, dataPart);
        }

        private void ReplaceParagraphText(XElement p, XNamespace a, string text) {
            // Remove existing runs; keep endParaRPr if present.
            var endPr = p.Element(a + "endParaRPr");
            p.Elements(a + "r").Remove();
            p.Elements(a + "br").Remove();
            InsertRunsWithBreaks(p, a, text, bold: false, italic: false);
        }

        private void ReplaceParagraphText(XElement p, XNamespace a, string text, bool bold, bool italic) {
            var endPr = p.Element(a + "endParaRPr");
            p.Elements(a + "r").Remove();
            p.Elements(a + "br").Remove();
            InsertRunsWithBreaks(p, a, text, bold, italic);
        }

        private void ReplaceParagraphText(XElement p, XNamespace a, string text, bool bold, bool italic, bool underline, string? colorHex, double? sizePt) {
            var endPr = p.Element(a + "endParaRPr");
            p.Elements(a + "r").Remove();
            p.Elements(a + "br").Remove();
            InsertRunsWithBreaks(p, a, text, bold, italic, underline, colorHex, sizePt);
        }

        private void InsertRunsWithBreaks(XElement p, XNamespace a, string text, bool bold, bool italic) {
            var lines = (text ?? string.Empty).Split('\n');
            XNamespace xml = XNamespace.Xml;
            for (int i = 0; i < lines.Length; i++) {
                if (i > 0) p.Add(new XElement(a + "br"));
                var rPr = new XElement(a + "rPr");
                if (bold) rPr.SetAttributeValue("b", 1);
                if (italic) rPr.SetAttributeValue("i", 1);
                var run = new XElement(a + "r",
                    rPr,
                    new XElement(a + "t", new XAttribute(xml + "space", "preserve"), lines[i]));
                // Insert before endParaRPr if present; else append to end
                var endPr = p.Element(a + "endParaRPr");
                if (endPr != null) endPr.AddBeforeSelf(run);
                else p.Add(run);
            }
        }

        private void InsertRunsWithBreaks(XElement p, XNamespace a, string text, bool bold, bool italic, bool underline, string? colorHex, double? sizePt) {
            var lines = (text ?? string.Empty).Split('\n');
            XNamespace xml = XNamespace.Xml;
            for (int i = 0; i < lines.Length; i++) {
                if (i > 0) p.Add(new XElement(a + "br"));
                var rPr = new XElement(a + "rPr");
                if (bold) rPr.SetAttributeValue("b", 1);
                if (italic) rPr.SetAttributeValue("i", 1);
                if (underline) rPr.SetAttributeValue("u", "sng");
                if (sizePt.HasValue) rPr.SetAttributeValue("sz", (int)Math.Round(sizePt.Value * 100));
                var run = new XElement(a + "r",
                    rPr,
                    new XElement(a + "t", new XAttribute(xml + "space", "preserve"), lines[i]));
                if (!string.IsNullOrWhiteSpace(colorHex)) {
                    var hex = colorHex!.Trim();
                    if (hex.StartsWith("#")) hex = hex.Substring(1);
                    if (hex.Length == 6) {
                        rPr.Add(new XElement(a + "solidFill",
                            new XElement(a + "srgbClr", new XAttribute("val", hex.ToUpper()))));
                    }
                }
                var endPr = p.Element(a + "endParaRPr");
                if (endPr != null) endPr.AddBeforeSelf(run);
                else p.Add(run);
            }
        }

        private SmartArtType? TryDetectType() {
            try {
                var rel = _drawing.Descendants<RelationshipIds>().FirstOrDefault();
                var main = _document._wordprocessingDocument.MainDocumentPart!;
                var layoutId = rel?.LayoutPart?.Value;
                if (string.IsNullOrEmpty(layoutId)) return null;
                var part = main.GetPartById(layoutId) as DiagramLayoutDefinitionPart;
                if (part?.LayoutDefinition == null) return null;
                var uid = part.LayoutDefinition.UniqueId?.Value ?? string.Empty;
                if (uid.EndsWith("/layout/cycle2")) return SmartArtType.Cycle;
                if (uid.EndsWith("/layout/default")) return SmartArtType.BasicProcess;
                return null;
            } catch { return null; }
        }

        private void SaveDiagramData(DiagramDataPart dataPart, XDocument xdoc) {
            using var ms = new MemoryStream();
            xdoc.Save(ms);
            ms.Position = 0;
            dataPart.FeedData(ms);
        }

        private XDocument LoadDiagramXDocument(DiagramDataPart dataPart) {
            // Prefer the part stream when available; fall back to DOM OuterXml when stream is empty.
            using (var stream = dataPart.GetStream(FileMode.Open, FileAccess.Read)) {
                if (stream != null && stream.Length > 0) {
                    return XDocument.Load(stream);
                }
            }
            if (dataPart.DataModelRoot != null) {
                var xml = dataPart.DataModelRoot.OuterXml;
                return XDocument.Parse(xml);
            }
            throw new InvalidOperationException("SmartArt data part is empty.");
        }

        private DiagramDataPart GetDiagramDataPart() {
            var rel = _drawing.Descendants<RelationshipIds>().FirstOrDefault();
            var dataId = rel?.DataPart?.Value;
            if (string.IsNullOrEmpty(dataId)) throw new InvalidOperationException("SmartArt data relationship not found.");
            var mainPart = _document._wordprocessingDocument.MainDocumentPart!;
            var part = mainPart.GetPartById(dataId);
            if (part is DiagramDataPart ddp) return ddp;
            throw new InvalidOperationException("DiagramDataPart not found for SmartArt.");
        }
    }
}
