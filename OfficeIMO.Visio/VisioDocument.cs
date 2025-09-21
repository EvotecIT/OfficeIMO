using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio document containing pages.
    /// </summary>
    public partial class VisioDocument {
        private readonly List<VisioPage> _pages = new();
        private bool _requestRecalcOnOpen;
        private string? _filePath;

        private const string DocumentRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/document";
        private const string DocumentContentType = "application/vnd.ms-visio.drawing.main+xml";
        private const string VisioNamespace = "http://schemas.microsoft.com/office/visio/2012/main";
        private const string ThemeRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/theme";
        private const string ThemeContentType = "application/vnd.ms-visio.theme+xml";
        private const string WindowsRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/windows";
        private const string WindowsContentType = "application/vnd.ms-visio.windows+xml";
        private const string PagesRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/pages";
        private const string PagesContentType = "application/vnd.ms-visio.pages+xml";
        private const string PageRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/page";
        private const string PageContentType = "application/vnd.ms-visio.page+xml";
        private const string MastersRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/masters";
        private const string MasterRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/master";

        /// <summary>
        /// Gets the collection of pages in the document.
        /// </summary>
        public IReadOnlyList<VisioPage> Pages => _pages;

        /// <summary>
        /// Gets or sets the theme applied to the document.
        /// </summary>
        public VisioTheme? Theme { get; set; }

        /// <summary>
        /// Gets or sets the title of the document.
        /// </summary>
        public string? Title { get; set; }

        /// <summary>
        /// Gets or sets the author of the document.
        /// </summary>
        public string? Author { get; set; }

        /// <summary>
        /// Adds a new page to the document.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        /// <param name="width">Page width.</param>
        /// <param name="height">Page height.</param>
        /// <param name="unit">Measurement unit for width and height.</param>
        /// <param name="id">Optional page identifier. If not specified, uses zero-based index.</param>
        public VisioPage AddPage(string name, double width = 8.26771653543307, double height = 11.69291338582677, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, int? id = null) {
            double widthInches = width.ToInches(unit);
            double heightInches = height.ToInches(unit);
            VisioPage page = new(name, widthInches, heightInches) { Id = id ?? _pages.Count };
            _pages.Add(page);
            return page;
        }

        /// <summary>
        /// Requests Visio to relayout and reroute connectors when the document is opened.
        /// </summary>
        public void RequestRecalcOnOpen() {
            _requestRecalcOnOpen = true;
        }

        /// <summary>
        /// Creates a new <see cref="VisioDocument"/> with the given save path.
        /// </summary>
        /// <param name="path">Path where the document will be saved.</param>
        public static VisioDocument Create(string path) {
            return new VisioDocument { _filePath = path };
        }
    }
}

