namespace OfficeIMO.Xps;

/// <summary>An editable native fixed document. Page references retain their attributes and link-target declarations.</summary>
public sealed class XpsFixedDocument {
    internal XpsDocument Owner { get; }
    internal XElement Markup { get; set; }
    internal XpsFixedDocument(XpsDocument owner, string partName, XElement markup) { Owner = owner; PartName = partName; Markup = markup; }
    /// <summary>Native package part name, without a leading slash.</summary>
    public string PartName { get; }
    /// <summary>Pages in this fixed document, including repeated references to shared native pages.</summary>
    public IReadOnlyList<XpsPage> Pages => Markup.Elements().Select(e => Owner.ResolvePage(PartName, e)).ToArray();
    /// <summary>Returns a detached copy of native fixed-document markup.</summary>
    public XElement GetMarkup() => new(Markup);
    /// <summary>Appends a new page, including when this fixed document was loaded from a package.</summary>
    public XpsPage AddPage(double width = 816, double height = 1056, string language = "en-US") => InsertPage(Markup.Elements().Count(), width, height, language);
    /// <summary>Inserts a new native page at the specified zero-based index.</summary>
    public XpsPage InsertPage(int index, double width = 816, double height = 1056, string language = "en-US") => InsertNewPage(index, width, height, language);
    internal XpsPage InsertNewPage(int index, double width, double height, string language, bool attach = false) {
        XpsDocument.CheckIndex(index, Markup.Elements().Count(), true);
        XpsPage.ValidatePageDimension(width); XpsPage.ValidatePageDimension(height);
        if (string.IsNullOrWhiteSpace(language)) throw new ArgumentException("A page language is required.", nameof(language));
        XNamespace ns = XpsPackage.Namespace(Owner.Format);
        string name = Owner.NewPartName(PartName.Substring(0, PartName.LastIndexOf('/') + 1) + "Pages/", ".fpage");
        var page = new XpsPage(Owner, name, new XElement(ns + "FixedPage", new XAttribute("Width", XpsPackage.N(width)), new XAttribute("Height", XpsPackage.N(height)), new XAttribute(XNamespace.Xml + "lang", language)));
        InsertPageCore(index, page, true, attach); return page;
    }
    /// <summary>Inserts another reference to a page owned by this package. The page and its relative resource base are shared.</summary>
    public void InsertPage(int index, XpsPage page) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        if (page.Document != Owner) throw new ArgumentException("The page belongs to another package.", nameof(page));
        XpsDocument.CheckIndex(index, Markup.Elements().Count(), true);
        InsertPageCore(index, page, false);
    }
    private void InsertPageCore(int index, XpsPage page, bool pending, bool attach = false) {
        var markup = new XElement(Markup);
        var reference = new XElement(markup.Name.Namespace + "PageContent", new XAttribute("Source", "/" + page.PartName), new XAttribute("Width", XpsPackage.N(page.Width)), new XAttribute("Height", XpsPackage.N(page.Height)));
        XpsDocument.InsertChild(markup, index, reference); Owner.CommitDocument(this, markup, pending ? page : null, attach);
    }
    /// <summary>Removes one page reference. The page part and associated resources remain in the package.</summary>
    public void RemovePageAt(int index) {
        XpsDocument.CheckIndex(index, Markup.Elements().Count(), false);
        var markup = new XElement(Markup); markup.Elements().ElementAt(index).Remove(); Owner.CommitDocument(this, markup);
    }
    /// <summary>Moves a page reference to its final zero-based index, preserving its native attributes and link targets.</summary>
    public void MovePage(int sourceIndex, int destinationIndex) {
        int count = Markup.Elements().Count(); XpsDocument.CheckIndex(sourceIndex, count, false); XpsDocument.CheckIndex(destinationIndex, count, false);
        var markup = new XElement(Markup); var reference = markup.Elements().ElementAt(sourceIndex); reference.Remove();
        XpsDocument.InsertChild(markup, destinationIndex, reference); Owner.CommitDocument(this, markup);
    }
    /// <summary>Moves one page reference to another fixed document in this package, retaining reference metadata and the page's resource base.</summary>
    /// <remarks>Repeated sequence references to either fixed document observe the change. Use MovePage for moves within one document.</remarks>
    public void MovePageTo(int sourceIndex, XpsFixedDocument destination, int destinationIndex) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (destination.Owner != Owner || destination == this) throw new ArgumentException("Choose a different fixed document in this package.", nameof(destination));
        XpsDocument.CheckIndex(sourceIndex, Markup.Elements().Count(), false);
        XpsDocument.CheckIndex(destinationIndex, destination.Markup.Elements().Count(), true);
        var source = new XElement(Markup); var target = new XElement(destination.Markup);
        var reference = source.Elements().ElementAt(sourceIndex);
        // The reference base changes; the native page part and all of its resource bases do not.
        reference.SetAttributeValue("Source", "/" + Owner.ResolvePage(PartName, reference).PartName);
        reference.Remove(); XpsDocument.InsertChild(target, destinationIndex, reference);
        Owner.CommitDocuments(new Dictionary<XpsFixedDocument, XElement> { [this] = source, [destination] = target });
    }

}
