using System;

namespace OfficeIMO.Visio.Fluent {
    /// <summary>
    /// Provides a fluent wrapper for <see cref="VisioDocument"/> allowing chained configuration.
    /// </summary>
    public class VisioFluentDocument {
        private readonly VisioDocument _document;

        internal VisioDocument Document => _document;

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioFluentDocument"/> class.
        /// </summary>
        /// <param name="document">The underlying <see cref="VisioDocument"/>.</param>
        public VisioFluentDocument(VisioDocument document) {
            _document = document;
        }

        /// <summary>
        /// Provides fluent access to document information.
        /// </summary>
        /// <param name="action">Action that receives an <see cref="InfoBuilder"/>.</param>
        public VisioFluentDocument Info(Action<InfoBuilder> action) {
            action(new InfoBuilder(this));
            return this;
        }

        /// <summary>
        /// Adds a page using a direct fluent style (no Add*/With* names) and configures it.
        /// Mirrors patterns from Markdown/Excel/PowerPoint fluent APIs.
        /// </summary>
        /// <param name="name">Page name.</param>
        /// <param name="configure">Configuration for shapes/connectors on the page.</param>
        public VisioFluentDocument Page(string name, Action<VisioFluentPage> configure) {
            var page = _document.AddPage(name);
            var builder = new VisioFluentPage(this, page);
            configure?.Invoke(builder);
            return this;
        }

        /// <summary>
        /// Adds a page with explicit size and configures it.
        /// </summary>
        public VisioFluentDocument Page(string name, double width, double height, VisioMeasurementUnit unit, Action<VisioFluentPage> configure) {
            var page = _document.AddPage(name, width, height, unit);
            var builder = new VisioFluentPage(this, page);
            configure?.Invoke(builder);
            return this;
        }

        // Removed obsolete AddPage overloads to keep the fluent API focused and consistent.

        /// <summary>
        /// Ends fluent configuration and returns the underlying document.
        /// </summary>
        public VisioDocument End() {
            return _document;
        }
    }
}
