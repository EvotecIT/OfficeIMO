using System.Globalization;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private const string NoPageRemovalNote = "Pages are kept; delete the page in Visio to remove it.";
    private const string HiddenLayerNote = "Shapes on a hidden layer can also belong to visible layers, so they are kept; remove them explicitly.";
    private const string HyperlinkNote = "The hyperlink belongs to the shape; edit or remove it on the shape to drop the path.";

    /// <summary>
    /// Finds metadata, personal information and hidden data that saving this drawing would keep, without changing it.
    /// Saving writes a new package from the drawing model, so package properties other than title and author, thumbnails and
    /// package-level Content Credentials of a loaded file are not carried into the saved file and are not reported.
    /// </summary>
    public OfficeHiddenDataReport InspectHiddenData() => CollectHiddenData(null);

    /// <summary>Removes the selected categories of metadata and hidden data, then reports what was found and what happened to each item. Save the drawing to keep the result.</summary>
    /// <param name="options">Categories to remove; defaults to <see cref="OfficeHiddenDataKind.Default"/>, which keeps hidden content and provenance.</param>
    public OfficeHiddenDataReport RemoveHiddenData(OfficeHiddenDataOptions? options = null) =>
        CollectHiddenData(options ?? new OfficeHiddenDataOptions());

    private OfficeHiddenDataReport CollectHiddenData(OfficeHiddenDataOptions? removal) {
        var collector = new OfficeHiddenDataCollector("Visio", removal);
        if (!string.IsNullOrEmpty(Properties.Author)) {
            collector.Add(OfficeHiddenDataKind.DocumentProperties, "/docProps/core.xml", "Author", Properties.Author, () => Properties.Author = null);
        }
        CollectVisioComments(collector);
        CollectVisioIdentities(collector);
        CollectVisioHiddenContent(collector);
        CollectVisioImages(collector);
        CollectVisioExternalReferences(collector);
        return collector.Build();
    }

    private void CollectVisioComments(OfficeHiddenDataCollector collector) {
        foreach (VisioPage page in _pages) {
            foreach (VisioComment comment in page.Comments.ToList()) {
                collector.Add(OfficeHiddenDataKind.Comments, PageLocation(page), "Comment by " + (comment.AuthorName ?? "unknown author"),
                    comment.Text, () => page.Comments.Remove(comment));
            }
        }
    }

    private void CollectVisioIdentities(OfficeHiddenDataCollector collector) {
        var authored = _pages.SelectMany(page => page.Comments.Select(comment => (Page: page, Comment: comment)))
            .Where(item => !string.IsNullOrEmpty(item.Comment.AuthorName) || !string.IsNullOrEmpty(item.Comment.AuthorResolutionId))
            .GroupBy(item => item.Comment.AuthorName ?? item.Comment.AuthorResolutionId!, StringComparer.Ordinal);
        foreach (var group in authored) {
            if (collector.Removal != null && group.Key == collector.Removal.AnonymousAuthor &&
                group.All(item => string.IsNullOrEmpty(item.Comment.AuthorResolutionId))) continue;
            string? resolution = group.Select(item => item.Comment.AuthorResolutionId).FirstOrDefault(id => !string.IsNullOrEmpty(id));
            collector.Add(OfficeHiddenDataKind.PersonalIdentities, string.Join(", ", group.Select(item => PageLocation(item.Page)).Distinct()),
                "Comment author on " + group.Count().ToString(CultureInfo.InvariantCulture) + " comments" +
                    (resolution == null ? string.Empty : ", with a resolution identity"), group.Key, () => {
                    var anonymous = new VisioCommentAuthor(collector.Removal!.AnonymousAuthor, collector.Removal.AnonymousInitials);
                    foreach (var item in group) item.Comment.SetAuthor(anonymous);
                }, removedOutcome: OfficeHiddenDataOutcome.Anonymized);
        }
    }

    private void CollectVisioHiddenContent(OfficeHiddenDataCollector collector) {
        foreach (VisioPage page in _pages) {
            if (page.UiVisibility == VisioPageUiVisibility.Hidden) {
                collector.Add(OfficeHiddenDataKind.HiddenContent, PageLocation(page), "Page hidden from the page tabs", null, null, NoPageRemovalNote);
            }
            foreach (VisioLayer layer in page.Layers.Where(layer => !layer.Visible)) {
                collector.Add(OfficeHiddenDataKind.HiddenContent, PageLocation(page), "Hidden layer '" + layer.Name + "'", null, null, HiddenLayerNote);
            }
        }
    }

    /// <summary>Embedded images on pages, shapes, connectors and masters, through the shared asset engine. A resource shared by several owners is processed once.</summary>
    private void CollectVisioImages(OfficeHiddenDataCollector collector) {
        var seen = new HashSet<VisioForeignResource>();
        foreach (VisioPage page in _pages) {
            IEnumerable<VisioForeignResource> resources = page.ForeignResources.Concat(ShapeForeignResources(page.Shapes))
                .Concat(page.Connectors.SelectMany(connector => connector.ForeignResources));
            Process(resources, PageLocation(page));
        }
        foreach (VisioMaster master in _registeredMasters) {
            Process(master.ForeignResources.Concat(ShapeForeignResources(new[] { master.Shape })), "Master '" + master.NameU + "'");
        }

        void Process(IEnumerable<VisioForeignResource> resources, string location) {
            foreach (VisioForeignResource resource in resources.ToList()) {
                if (!resource.RelationshipType.EndsWith("/image", StringComparison.Ordinal) || !seen.Add(resource)) continue;
                byte[]? replacement = OfficeHiddenDataAssets.ProcessImage(resource.Bytes, location, collector);
                if (replacement != null) resource.Bytes = replacement;
            }
        }
    }

    /// <summary>Hyperlinks that point to files or network paths. They are part of the drawing, so they are reported and kept.</summary>
    private void CollectVisioExternalReferences(OfficeHiddenDataCollector collector) {
        foreach (VisioPage page in _pages) {
            foreach (VisioShape shape in AllShapes(page.Shapes)) Report(shape.Hyperlinks, PageLocation(page) + ", shape '" + (shape.NameU ?? shape.Id) + "'");
            foreach (VisioConnector connector in page.Connectors) Report(connector.Hyperlinks, PageLocation(page) + ", connector '" + connector.Id + "'");
        }

        void Report(IEnumerable<VisioHyperlink> hyperlinks, string location) {
            foreach (VisioHyperlink hyperlink in hyperlinks) {
                if (!IsFileAddress(hyperlink.Address)) continue;
                collector.Add(OfficeHiddenDataKind.ExternalReferences, location, "Hyperlink to a file or network path", hyperlink.Address, null, HyperlinkNote);
            }
        }
    }

    private static IEnumerable<VisioShape> AllShapes(IEnumerable<VisioShape> shapes) {
        foreach (VisioShape shape in shapes) {
            yield return shape;
            foreach (VisioShape child in AllShapes(shape.Children)) yield return child;
        }
    }

    private static bool IsFileAddress(string? address) {
        if (string.IsNullOrWhiteSpace(address)) return false;
        string value = address!.Trim();
        if (value.StartsWith("file:", StringComparison.OrdinalIgnoreCase) || value.StartsWith("\\\\", StringComparison.Ordinal) ||
            value.StartsWith("/", StringComparison.Ordinal) || value.StartsWith(".", StringComparison.Ordinal)) return true;
        if (value.Length >= 3 && char.IsLetter(value[0]) && value[1] == ':' && (value[2] == '\\' || value[2] == '/')) return true;
        int colon = value.IndexOf(':');
        return colon < 0;
    }

    private static string PageLocation(VisioPage page) => "Page '" + page.Name + "'";
}
