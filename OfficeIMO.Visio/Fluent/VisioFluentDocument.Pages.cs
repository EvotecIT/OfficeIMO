using System;
using System.Linq;

namespace OfficeIMO.Visio.Fluent {
    public partial class VisioFluentDocument {
        /// <summary>
        /// Configures the first page in an existing document.
        /// </summary>
        /// <param name="configure">Configuration to apply to the first page.</param>
        public VisioFluentDocument FirstPage(Action<VisioFluentPage> configure) {
            if (_document.Pages.Count == 0) {
                throw new InvalidOperationException("The document does not contain any pages.");
            }

            return Page(_document.Pages[0], configure);
        }

        /// <summary>
        /// Configures an existing page by zero-based index without adding a new page.
        /// </summary>
        /// <param name="pageIndex">Zero-based page index.</param>
        /// <param name="configure">Configuration to apply to the page.</param>
        public VisioFluentDocument ExistingPage(int pageIndex, Action<VisioFluentPage> configure) {
            if (pageIndex < 0 || pageIndex >= _document.Pages.Count) {
                throw new ArgumentOutOfRangeException(nameof(pageIndex), "Page index is outside the document page collection.");
            }

            return Page(_document.Pages[pageIndex], configure);
        }

        /// <summary>
        /// Configures an existing page by name without adding a duplicate page.
        /// </summary>
        /// <param name="name">Page name to find.</param>
        /// <param name="configure">Configuration to apply to the page.</param>
        /// <param name="comparison">String comparison used when matching page names.</param>
        public VisioFluentDocument ExistingPage(string name, Action<VisioFluentPage> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Page name cannot be null or whitespace.", nameof(name));
            }

            VisioPage? page = _document.Pages.FirstOrDefault(candidate => string.Equals(candidate.Name, name, comparison));
            if (page == null) {
                throw new InvalidOperationException($"Page '{name}' was not found.");
            }

            return Page(page, configure);
        }

        /// <summary>
        /// Configures an existing page instance that belongs to this document.
        /// </summary>
        /// <param name="page">Existing page instance.</param>
        /// <param name="configure">Configuration to apply to the page.</param>
        public VisioFluentDocument Page(VisioPage page, Action<VisioFluentPage> configure) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (!_document.Pages.Any(candidate => ReferenceEquals(candidate, page))) {
                throw new InvalidOperationException("The page does not belong to this document.");
            }

            ConfigurePage(page, configure);
            return this;
        }

        /// <summary>
        /// Configures an existing page by name or adds it when it is missing.
        /// </summary>
        /// <param name="name">Page name to find or create.</param>
        /// <param name="configure">Configuration to apply to the page.</param>
        /// <param name="comparison">String comparison used when matching page names.</param>
        public VisioFluentDocument PageOrAdd(string name, Action<VisioFluentPage> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Page name cannot be null or whitespace.", nameof(name));
            }

            VisioPage page = _document.Pages.FirstOrDefault(candidate => string.Equals(candidate.Name, name, comparison))
                ?? _document.AddPage(name);
            ConfigurePage(page, configure);
            return this;
        }

        /// <summary>
        /// Configures an existing page by name or adds it with explicit size when it is missing.
        /// </summary>
        /// <param name="name">Page name to find or create.</param>
        /// <param name="width">Page width to use when creating a missing page.</param>
        /// <param name="height">Page height to use when creating a missing page.</param>
        /// <param name="unit">Measurement unit for a newly created page.</param>
        /// <param name="configure">Configuration to apply to the page.</param>
        /// <param name="comparison">String comparison used when matching page names.</param>
        public VisioFluentDocument PageOrAdd(
            string name,
            double width,
            double height,
            VisioMeasurementUnit unit,
            Action<VisioFluentPage> configure,
            StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Page name cannot be null or whitespace.", nameof(name));
            }

            VisioPage page = _document.Pages.FirstOrDefault(candidate => string.Equals(candidate.Name, name, comparison))
                ?? _document.AddPage(name, width, height, unit);
            ConfigurePage(page, configure);
            return this;
        }

        private void ConfigurePage(VisioPage page, Action<VisioFluentPage> configure) {
            if (configure == null) {
                throw new ArgumentNullException(nameof(configure));
            }

            VisioFluentPage builder = new(this, page);
            configure(builder);
            builder.RebuildShapeIndex();
        }
    }
}
