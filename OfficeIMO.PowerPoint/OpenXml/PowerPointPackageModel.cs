using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint.OpenXml;

/// <summary>
/// A presentation package edited through the Open XML layer instead of the Open XML SDK: parts are trees parsed on
/// first use, slides are edited through generated PresentationML views, and the package is saved through the shared
/// ZIP writer. Untouched parts keep their bytes.
/// </summary>
/// <remarks>This is the model prototype that the engine may move onto; nothing in the public API uses it yet.</remarks>
internal sealed class PowerPointPackageModel : IDisposable {
    internal const string PresentationPartName = "/ppt/presentation.xml";
    internal const string PresentationContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    internal const string SlideContentType = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
    internal const string SlideLayoutContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
    internal const string SlideMasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
    internal const string ThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";
    internal const string PresentationPropertiesContentType = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";
    internal const string TableStylesContentType = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";
    internal const string CorePropertiesContentType = "application/vnd.openxmlformats-package.core-properties+xml";
    internal const string ExtendedPropertiesContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";

    private const string RelationshipTypes = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix;
    private const string PresentationMLUri = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private const string DrawingMLUri = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string RelationshipsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const uint FirstSlideId = 256;
    // ST_SlideId: 256 up to, but not including, 2147483648.
    private const uint LastSlideId = 2147483647;

    private static readonly Lazy<byte[]> Theme = new(() => LoadResource("OfficeIMO.PowerPoint.Resources.theme1.xml"));
    private static readonly Lazy<byte[]> TableStyles = new(() => LoadResource("OfficeIMO.PowerPoint.Resources.tableStyles.xml"));

    private readonly OfficeOpcEditablePackage _package;
    private readonly string _presentationPart;
    private readonly OfficeXmlTree _presentation;
    private string? _defaultLayoutPart;
    private string? _defaultLayoutTarget;
    private int _nextSlidePart;
    // A Strict deck would need Strict markup and relationship types for every new part, which this model does not write.
    private readonly bool _strict;

    private PowerPointPackageModel(OfficeOpcEditablePackage package, string presentationPart, bool strict = false) {
        _package = package;
        _presentationPart = presentationPart;
        _presentation = package.GetXml(presentationPart);
        _strict = strict;
        _nextSlidePart = 1;
        foreach (string name in package.PartNames) {
            if (!name.StartsWith("/ppt/slides/slide", StringComparison.OrdinalIgnoreCase) || !name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;
            if (int.TryParse(name.Substring(17, name.Length - 21), NumberStyles.None, CultureInfo.InvariantCulture, out int number) && number >= _nextSlidePart) {
                _nextSlidePart = number + 1;
            }
        }
    }

    internal PPresentationView Presentation => new(_presentation, _presentation.Root);

    internal int SlideCount => Presentation.SldIdLst?.SldIdList.Count ?? 0;

    /// <summary>Opens a presentation from a readable, seekable stream, which must stay open until the model is disposed.</summary>
    internal static PowerPointPackageModel Open(Stream stream) {
        OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(stream, leaveOpen: true);
        try {
            string? presentation = null;
            bool strict = false;
            OfficeXmlTree rels = package.GetRelationships(null);
            for (int node = rels.FindChild(rels.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                node = rels.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                string? type = rels.GetAttribute(node, OfficeXmlNamespaces.None, "Type");
                if (type != null && OfficeXmlNamespaces.NormalizeRelationshipType(type) == RelationshipTypes + "officeDocument") {
                    presentation = OfficeOpcPartNames.ResolveRelationshipTarget("/_rels/.rels", rels.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                    strict = !string.Equals(type, RelationshipTypes + "officeDocument", StringComparison.Ordinal);
                    break;
                }
            }
            if (presentation == null || !package.PartExists(presentation)) throw new InvalidDataException("The package has no presentation part.");
            return new PowerPointPackageModel(package, presentation, strict);
        } catch {
            package.Dispose();
            throw;
        }
    }

    /// <summary>Creates a 16:9 presentation with one master, a Title and Content layout, the Office theme and no slides.</summary>
    internal static PowerPointPackageModel Create() {
        OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Create();
        try {
            package.AddRelationship(null, RelationshipTypes + "officeDocument", "ppt/presentation.xml");
            package.AddRelationship(null, "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
            package.AddRelationship(null, RelationshipTypes + "extended-properties", "docProps/app.xml");
            CreateDocumentProperties(package);

            OfficeXmlTree presentationTree = package.CreateXmlPart(PresentationPartName, PresentationContentType, OfficeXmlNamespaces.PresentationML, "presentation");
            DeclareSlideNamespaces(presentationTree);
            var presentation = new PPresentationView(presentationTree, presentationTree.Root) { SaveSubsetFonts = true };
            string masterRelationship = package.AddRelationship(PresentationPartName, RelationshipTypes + "slideMaster", "slideMasters/slideMaster1.xml");
            PSlideMasterIdListEntryView masterId = presentation.GetOrAddSldMasterIdLst().AddSldMasterId();
            masterId.Id = 2147483648;
            masterId.RelationshipId = masterRelationship;
            presentation.GetOrAddSldIdLst();
            PSlideSizeView size = presentation.GetOrAddSldSz();
            size.Cx = 12192000;
            size.Cy = 6858000;
            APositiveSize2DView notes = presentation.GetOrAddNotesSz();
            notes.Cx = 6858000;
            notes.Cy = 9144000;

            package.AddRelationship(PresentationPartName, RelationshipTypes + "theme", "theme/theme1.xml");
            package.CreateBinaryPart("/ppt/theme/theme1.xml", ThemeContentType, Theme.Value);
            package.AddRelationship(PresentationPartName, RelationshipTypes + "presProps", "presProps.xml");
            OfficeXmlTree presProps = package.CreateXmlPart("/ppt/presProps.xml", PresentationPropertiesContentType, OfficeXmlNamespaces.PresentationML, "presentationPr", 4);
            DeclareSlideNamespaces(presProps);
            package.AddRelationship(PresentationPartName, RelationshipTypes + "tableStyles", "tableStyles.xml");
            package.CreateBinaryPart("/ppt/tableStyles.xml", TableStylesContentType, TableStyles.Value);

            CreateMasterAndLayout(package);
            return new PowerPointPackageModel(package, PresentationPartName);
        } catch {
            package.Dispose();
            throw;
        }
    }

    /// <summary>Adds a slide on the first layout of the first master and returns it.</summary>
    internal PowerPointModelSlide AddSlide() {
        if (_strict) throw new NotSupportedException("Adding slides to a Strict Open XML presentation is not supported yet.");
        uint slideId = NextSlideId();
        string layout = _defaultLayoutPart ??= FindDefaultLayout();
        // Slides share a folder, so every new slide reaches the layout by the same relative target.
        string layoutTarget = _defaultLayoutTarget ??= RelativeTarget("/ppt/slides/slide1.xml", layout);
        int number = _nextSlidePart++;
        string partName = "/ppt/slides/slide" + number.ToString(CultureInfo.InvariantCulture) + ".xml";
        OfficeXmlTree tree = _package.CreateXmlPart(partName, SlideContentType, OfficeXmlNamespaces.PresentationML, "sld", 64);
        DeclareSlideNamespaces(tree);
        var slide = new PSlideView(tree, tree.Root);
        PGroupShapeView shapes = slide.GetOrAddCSld().GetOrAddSpTree();
        PGroupShapeNonVisualView group = shapes.GetOrAddNvGrpSpPr();
        ANonVisualDrawingPropsView groupProperties = group.GetOrAddCNvPr();
        groupProperties.Id = 1;
        groupProperties.Name = string.Empty;
        group.GetOrAddCNvGrpSpPr();
        group.GetOrAddNvPr();
        AGroupTransform2DView transform = shapes.GetOrAddGrpSpPr().GetOrAddXfrm();
        SetPoint(transform.GetOrAddOff(), 0, 0);
        SetSize(transform.GetOrAddExt(), 0, 0);
        SetPoint(transform.GetOrAddChOff(), 0, 0);
        SetSize(transform.GetOrAddChExt(), 0, 0);
        slide.GetOrAddClrMapOvr().GetOrAddMasterClrMapping();

        _package.AddRelationship(partName, RelationshipTypes + "slideLayout", layoutTarget);
        string relationship = _package.AddRelationship(_presentationPart, RelationshipTypes + "slide", RelativeTarget(_presentationPart, partName));
        PSlideIdListEntryView entry = Presentation.GetOrAddSldIdLst().AddSldId();
        entry.Id = slideId;
        entry.RelationshipId = relationship;
        return new PowerPointModelSlide(slide);
    }

    // Read from the slide list each time, since callers can change it through the Presentation view.
    private uint NextSlideId() {
        uint highest = FirstSlideId - 1;
        foreach (uint id in SlideIds()) {
            if (id > highest) highest = id;
        }
        if (highest < LastSlideId) return highest + 1;
        return PowerPointModelSlide.LowestUnused(SlideIds(), FirstSlideId, LastSlideId) ?? throw new InvalidOperationException("The presentation has no free slide id.");
    }

    private IEnumerable<uint> SlideIds() {
        foreach (int node in Presentation.SldIdLst?.SldIdList ?? default) {
            if (new PSlideIdListEntryView(_presentation, node).Id is { } id) yield return id;
        }
    }

    /// <summary>The slide at a zero-based position in the slide list.</summary>
    internal PowerPointModelSlide GetSlide(int index) {
        OfficeXmlTree tree = _package.GetXml(GetSlidePartName(index));
        return new PowerPointModelSlide(new PSlideView(tree, tree.Root));
    }

    /// <summary>The part name of the slide at a zero-based position in the slide list.</summary>
    internal string GetSlidePartName(int index) {
        int position = 0;
        foreach (int node in Presentation.SldIdLst?.SldIdList ?? default) {
            if (position++ != index) continue;
            string id = new PSlideIdListEntryView(_presentation, node).RelationshipId ?? throw new InvalidDataException("A slide entry has no relationship id.");
            return _package.GetRelationshipTarget(_presentationPart, id) ?? throw new InvalidDataException("A slide relationship has no target part.");
        }
        throw new ArgumentOutOfRangeException(nameof(index));
    }

    internal void Save(Stream destination) => _package.Save(destination);

    public void Dispose() => _package.Dispose();

    private string FindDefaultLayout() {
        PSlideMasterIdListEntryView? master = null;
        foreach (int node in Presentation.SldMasterIdLst?.SldMasterIdList ?? default) {
            master = new PSlideMasterIdListEntryView(_presentation, node);
            break;
        }
        string masterPart = master?.RelationshipId is { } masterId ? _package.GetRelationshipTarget(_presentationPart, masterId) ?? string.Empty : string.Empty;
        if (masterPart.Length == 0) throw new InvalidDataException("The presentation has no slide master.");
        OfficeXmlTree masterTree = _package.GetXml(masterPart);
        foreach (int node in new PSlideMasterView(masterTree, masterTree.Root).SldLayoutIdLst?.SldLayoutIdList ?? default) {
            string? layoutId = new PSlideLayoutIdListEntryView(masterTree, node).RelationshipId;
            if (layoutId != null && _package.GetRelationshipTarget(masterPart, layoutId) is { } layout) return layout;
        }
        throw new InvalidDataException("The slide master has no layouts.");
    }

    private static void CreateMasterAndLayout(OfficeOpcEditablePackage package) {
        const string masterPart = "/ppt/slideMasters/slideMaster1.xml";
        const string layoutPart = "/ppt/slideLayouts/slideLayout1.xml";
        OfficeXmlTree masterTree = package.CreateXmlPart(masterPart, SlideMasterContentType, OfficeXmlNamespaces.PresentationML, "sldMaster", 64);
        DeclareSlideNamespaces(masterTree);
        var master = new PSlideMasterView(masterTree, masterTree.Root);
        PGroupShapeView masterShapes = master.GetOrAddCSld().GetOrAddSpTree();
        AddGroupProperties(masterShapes);
        AddPlaceholder(masterShapes, 2, "Title Placeholder 1", "title", null, 838200, 365125, 10515600, 1325563);
        AddPlaceholder(masterShapes, 3, "Text Placeholder 2", "body", 1, 838200, 1825625, 10515600, 4351338);
        AColorMappingView colors = master.GetOrAddClrMap();
        colors.Bg1 = "lt1";
        colors.Tx1 = "dk1";
        colors.Bg2 = "lt2";
        colors.Tx2 = "dk2";
        colors.Accent1 = "accent1";
        colors.Accent2 = "accent2";
        colors.Accent3 = "accent3";
        colors.Accent4 = "accent4";
        colors.Accent5 = "accent5";
        colors.Accent6 = "accent6";
        colors.Hlink = "hlink";
        colors.FolHlink = "folHlink";
        string layoutRelationship = package.AddRelationship(masterPart, RelationshipTypes + "slideLayout", "../slideLayouts/slideLayout1.xml");
        PSlideLayoutIdListEntryView layoutId = master.GetOrAddSldLayoutIdLst().AddSldLayoutId();
        layoutId.Id = 2147483649;
        layoutId.RelationshipId = layoutRelationship;
        package.AddRelationship(masterPart, RelationshipTypes + "theme", "../theme/theme1.xml");

        OfficeXmlTree layoutTree = package.CreateXmlPart(layoutPart, SlideLayoutContentType, OfficeXmlNamespaces.PresentationML, "sldLayout", 48);
        DeclareSlideNamespaces(layoutTree);
        var layout = new PSlideLayoutView(layoutTree, layoutTree.Root) { Type = "obj", Preserve = true };
        PCommonSlideDataView layoutData = layout.GetOrAddCSld();
        layoutData.Name = "Title and Content";
        PGroupShapeView layoutShapes = layoutData.GetOrAddSpTree();
        AddGroupProperties(layoutShapes);
        AddPlaceholder(layoutShapes, 2, "Title 1", "title", null, null, null, null, null);
        AddPlaceholder(layoutShapes, 3, "Content Placeholder 2", null, 1, null, null, null, null);
        layout.GetOrAddClrMapOvr().GetOrAddMasterClrMapping();
        package.AddRelationship(layoutPart, RelationshipTypes + "slideMaster", "../slideMasters/slideMaster1.xml");
    }

    private static void AddGroupProperties(PGroupShapeView shapes) {
        PGroupShapeNonVisualView group = shapes.GetOrAddNvGrpSpPr();
        ANonVisualDrawingPropsView properties = group.GetOrAddCNvPr();
        properties.Id = 1;
        properties.Name = string.Empty;
        group.GetOrAddCNvGrpSpPr();
        group.GetOrAddNvPr();
        AGroupTransform2DView transform = shapes.GetOrAddGrpSpPr().GetOrAddXfrm();
        SetPoint(transform.GetOrAddOff(), 0, 0);
        SetSize(transform.GetOrAddExt(), 0, 0);
        SetPoint(transform.GetOrAddChOff(), 0, 0);
        SetSize(transform.GetOrAddChExt(), 0, 0);
    }

    private static void AddPlaceholder(PGroupShapeView shapes, uint id, string name, string? type, uint? index, long? x, long? y, long? cx, long? cy) {
        PShapeView shape = shapes.AddSp();
        PShapeNonVisualView nonVisual = shape.GetOrAddNvSpPr();
        ANonVisualDrawingPropsView properties = nonVisual.GetOrAddCNvPr();
        properties.Id = id;
        properties.Name = name;
        nonVisual.GetOrAddCNvSpPr();
        PPlaceholderView placeholder = nonVisual.GetOrAddNvPr().GetOrAddPh();
        placeholder.Type = type;
        placeholder.Idx = index;
        AShapePropertiesView shapeProperties = shape.GetOrAddSpPr();
        if (x.HasValue && y.HasValue && cx.HasValue && cy.HasValue) {
            ATransform2DView transform = shapeProperties.GetOrAddXfrm();
            SetPoint(transform.GetOrAddOff(), x.Value, y.Value);
            SetSize(transform.GetOrAddExt(), cx.Value, cy.Value);
            shapeProperties.GetOrAddPrstGeom().Prst = "rect";
        }
        ATextBodyView text = shape.GetOrAddTxBody();
        text.GetOrAddBodyPr();
        text.GetOrAddLstStyle();
        text.AddP();
    }

    private static void CreateDocumentProperties(OfficeOpcEditablePackage package) {
        OfficeXmlTree core = package.CreateXmlPart("/docProps/core.xml", CorePropertiesContentType, OfficeXmlNamespaces.CoreProperties, "coreProperties", 8);
        core.DeclareNamespace(core.Root, "cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
        core.DeclareNamespace(core.Root, "dc", "http://purl.org/dc/elements/1.1/");
        core.DeclareNamespace(core.Root, "dcterms", "http://purl.org/dc/terms/");
        core.DeclareNamespace(core.Root, "xsi", "http://www.w3.org/2001/XMLSchema-instance");
        int creator = core.CreateElement(OfficeXmlNamespaces.DublinCore, "creator");
        core.Append(core.Root, creator);
        core.Append(creator, core.CreateText("OfficeIMO"));
        int created = core.CreateElement(OfficeXmlNamespaces.DublinCoreTerms, "created");
        core.SetAttribute(created, core.Namespaces.GetId("http://www.w3.org/2001/XMLSchema-instance"), "type", "dcterms:W3CDTF");
        core.Append(core.Root, created);
        core.Append(created, core.CreateText(DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture)));

        OfficeXmlTree app = package.CreateXmlPart("/docProps/app.xml", ExtendedPropertiesContentType, OfficeXmlNamespaces.ExtendedProperties, "Properties", 4);
        app.DeclareNamespace(app.Root, string.Empty, "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
        int application = app.CreateElement(OfficeXmlNamespaces.ExtendedProperties, "Application");
        app.Append(app.Root, application);
        app.Append(application, app.CreateText("OfficeIMO"));
    }

    private static void DeclareSlideNamespaces(OfficeXmlTree tree) {
        tree.DeclareNamespace(tree.Root, "a", DrawingMLUri);
        tree.DeclareNamespace(tree.Root, "r", RelationshipsUri);
        tree.DeclareNamespace(tree.Root, "p", PresentationMLUri);
    }

    internal static void SetPoint(APoint2DView point, long x, long y) {
        point.X = x.ToString(CultureInfo.InvariantCulture);
        point.Y = y.ToString(CultureInfo.InvariantCulture);
    }

    internal static void SetSize(APositiveSize2DView size, long cx, long cy) {
        size.Cx = cx;
        size.Cy = cy;
    }

    // Both parts sit under /ppt, so the target is relative to the source part's folder.
    private static string RelativeTarget(string sourcePart, string targetPart) {
        string sourceFolder = sourcePart.Substring(0, sourcePart.LastIndexOf('/') + 1);
        if (targetPart.StartsWith(sourceFolder, StringComparison.Ordinal)) return targetPart.Substring(sourceFolder.Length);
        string parent = sourceFolder.Substring(0, sourceFolder.LastIndexOf('/', sourceFolder.Length - 2) + 1);
        if (targetPart.StartsWith(parent, StringComparison.Ordinal)) return "../" + targetPart.Substring(parent.Length);
        return targetPart;
    }

    private static byte[] LoadResource(string name) {
        using Stream stream = typeof(PowerPointPackageModel).Assembly.GetManifestResourceStream(name)
            ?? throw new InvalidOperationException("Missing embedded resource '" + name + "'.");
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }
}

/// <summary>One slide of a <see cref="PowerPointPackageModel"/>.</summary>
internal readonly struct PowerPointModelSlide {
    internal PowerPointModelSlide(PSlideView slide) => View = slide;

    internal PSlideView View { get; }

    internal PGroupShapeView Shapes => View.GetOrAddCSld().GetOrAddSpTree();

    /// <summary>Adds a title placeholder with one run of text.</summary>
    internal PShapeView AddTitle(string text) {
        PShapeView shape = AddShape("Title", out _);
        shape.GetOrAddNvSpPr().GetOrAddNvPr().GetOrAddPh().Type = "title";
        shape.GetOrAddSpPr();
        ATextBodyView body = shape.GetOrAddTxBody();
        body.GetOrAddBodyPr();
        body.GetOrAddLstStyle();
        AddParagraph(body, text);
        return shape;
    }

    /// <summary>Adds a rectangular text box with one paragraph per string.</summary>
    internal PShapeView AddTextBox(long x, long y, long cx, long cy, params string[] paragraphs) {
        PShapeView shape = AddShape("TextBox", out PShapeNonVisualView nonVisual);
        nonVisual.GetOrAddCNvSpPr().TxBox = true;
        nonVisual.GetOrAddNvPr();
        AShapePropertiesView properties = shape.GetOrAddSpPr();
        ATransform2DView transform = properties.GetOrAddXfrm();
        PowerPointPackageModel.SetPoint(transform.GetOrAddOff(), x, y);
        PowerPointPackageModel.SetSize(transform.GetOrAddExt(), cx, cy);
        properties.GetOrAddPrstGeom().GetOrAddAvLst();
        properties.GetOrAddPrstGeom().Prst = "rect";
        properties.GetOrAddNoFill();
        ATextBodyView body = shape.GetOrAddTxBody();
        ATextBodyPropertiesView bodyProperties = body.GetOrAddBodyPr();
        bodyProperties.Wrap = "square";
        bodyProperties.RtlCol = false;
        body.GetOrAddLstStyle();
        foreach (string paragraph in paragraphs) AddParagraph(body, paragraph);
        // A text body holds at least one paragraph, even when it is empty.
        if (paragraphs.Length == 0) body.AddP();
        return shape;
    }

    /// <summary>The title placeholder, or null.</summary>
    internal PShapeView? FindTitle() {
        PGroupShapeView shapes = Shapes;
        foreach (int node in shapes.SpList) {
            var shape = new PShapeView(shapes.Tree, node);
            string? type = shape.NvSpPr?.NvPr?.Ph?.Type;
            if (type == "title" || type == "ctrTitle") return shape;
        }
        return null;
    }

    /// <summary>Replaces a shape's text with one paragraph, keeping its body properties and list style.</summary>
    internal static void SetText(PShapeView shape, string text) {
        ATextBodyView body = shape.GetOrAddTxBody();
        // A shape that had no text gets the body properties every text body needs.
        body.GetOrAddBodyPr();
        OfficeXmlTree tree = body.Tree;
        int paragraph;
        while ((paragraph = tree.FindChild(body.Node, OfficeXmlNamespaces.DrawingML, "p")) != OfficeXmlTree.NoNode) tree.Remove(paragraph);
        AddParagraph(body, text);
    }

    // Office names shapes by kind and number, counting from the id it gives them.
    private PShapeView AddShape(string kind, out PShapeNonVisualView nonVisual) {
        uint id = NextShapeId();
        PShapeView shape = Shapes.AddSp();
        nonVisual = shape.GetOrAddNvSpPr();
        ANonVisualDrawingPropsView properties = nonVisual.GetOrAddCNvPr();
        properties.Id = id;
        properties.Name = kind + " " + (id - 1).ToString(CultureInfo.InvariantCulture);
        nonVisual.GetOrAddCNvSpPr();
        return shape;
    }

    // Drawing ids are unique across the slide: shapes, pictures, frames, connectors and groups, nested ones included.
    // A new one goes above the highest; when that is the largest id there is, the lowest unused one is taken.
    private uint NextShapeId() {
        uint highest = 1;
        foreach (uint id in DrawingIds()) {
            if (id > highest) highest = id;
        }
        if (highest < uint.MaxValue) return highest + 1;
        return LowestUnused(DrawingIds(), 2, uint.MaxValue) ?? throw new InvalidOperationException("The slide has no free drawing id.");
    }

    private IEnumerable<uint> DrawingIds() {
        PGroupShapeView shapes = Shapes;
        OfficeXmlTree tree = shapes.Tree;
        int root = shapes.Node;
        int node = root;
        while (true) {
            if (tree.IsElement(node, OfficeXmlNamespaces.PresentationML, "cNvPr")
                && uint.TryParse(tree.GetAttribute(node, OfficeXmlNamespaces.None, "id"), NumberStyles.None, CultureInfo.InvariantCulture, out uint id)) {
                yield return id;
            }
            int child = tree.KindOf(node) == OfficeXmlTreeNodeKind.Element ? tree.FirstChildOf(node) : OfficeXmlTree.NoNode;
            if (child != OfficeXmlTree.NoNode) {
                node = child;
                continue;
            }
            while (node != root && tree.NextSiblingOf(node) == OfficeXmlTree.NoNode) node = tree.ParentOf(node);
            if (node == root) yield break;
            node = tree.NextSiblingOf(node);
        }
    }

    internal static uint? LowestUnused(IEnumerable<uint> ids, uint first, uint last) {
        var used = new HashSet<uint>(ids);
        for (uint candidate = first; ; candidate++) {
            if (!used.Contains(candidate)) return candidate;
            if (candidate == last) return null;
        }
    }

    private static void AddParagraph(ATextBodyView body, string text) {
        ARegularTextRunView run = body.AddP().AddR();
        run.GetOrAddRPr().Lang = "en-US";
        run.T = text;
    }
}
