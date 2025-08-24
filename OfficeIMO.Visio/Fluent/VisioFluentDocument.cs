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
        /// Adds a page and returns the fluent document for chaining.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        /// <param name="width">Page width.</param>
        /// <param name="height">Page height.</param>
        /// <param name="unit">Measurement unit for width and height.</param>
        /// <param name="page">The created page.</param>
        public VisioFluentDocument AddPage(string name, double width, double height, VisioMeasurementUnit unit, out VisioPage page) {
            page = _document.AddPage(name, width, height, unit);
            return this;
        }

        /// <summary>
        /// Adds a page with the specified width and height.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        /// <param name="width">Page width.</param>
        /// <param name="height">Page height.</param>
        /// <param name="page">The created page.</param>
        /// <remarks>This overload is obsolete. Use <see cref="AddPage(string,double,double,VisioMeasurementUnit,out VisioPage)"/> instead.</remarks>
        [System.Obsolete("Use AddPage with width, height and unit parameters")]
        public VisioFluentDocument AddPage(string name, double width, double height, out VisioPage page) {
            page = _document.AddPage(name, width, height);
            return this;
        }

        /// <summary>
        /// Adds a page with default dimensions.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        /// <param name="page">The created page.</param>
        /// <remarks>This overload is obsolete. Use <see cref="AddPage(string,double,double,VisioMeasurementUnit,out VisioPage)"/> instead.</remarks>
        [System.Obsolete("Use AddPage with width, height and unit parameters")]
        public VisioFluentDocument AddPage(string name, out VisioPage page) {
            page = _document.AddPage(name);
            return this;
        }

        /// <summary>
        /// Ends fluent configuration and returns the underlying document.
        /// </summary>
        public VisioDocument End() {
            return _document;
        }
    }
}
