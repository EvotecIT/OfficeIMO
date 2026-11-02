using System;

namespace OfficeIMO.Visio.Fluent {
    /// <summary>
    /// Provides a fluent wrapper for <see cref="VisioDocument"/> allowing chained configuration.
    /// </summary>
    public partial class VisioFluentDocument {
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

        /// <summary>Adds a page and optionally configures it.</summary>
        public VisioFluentDocument AddPage(string name, OfficeLength? width = null, OfficeLength? height = null,
            VisioMeasurementUnit scaleUnit = VisioMeasurementUnit.Inches, int? id = null, Action<VisioFluentPage>? configure = null) {
            ConfigurePage(_document.AddPage(name, width, height, scaleUnit, id), configure);
            return this;
        }

        /// <summary>Adds a default-sized page and configures it.</summary>
        public VisioFluentDocument AddPage(string name, Action<VisioFluentPage> configure) => AddPage(name, configure: configure, width: null);

        /// <summary>Adds a background page and optionally configures it.</summary>
        public VisioFluentDocument AddBackgroundPage(string name, OfficeLength? width = null, OfficeLength? height = null,
            VisioMeasurementUnit scaleUnit = VisioMeasurementUnit.Inches, int? id = null, Action<VisioFluentPage>? configure = null) {
            ConfigurePage(_document.AddBackgroundPage(name, width, height, scaleUnit, id), configure);
            return this;
        }

        /// <summary>Adds a default-sized background page and configures it.</summary>
        public VisioFluentDocument AddBackgroundPage(string name, Action<VisioFluentPage> configure) => AddBackgroundPage(name, configure: configure, width: null);

        /// <summary>
        /// Applies high-level deterministic cleanup to all foreground pages.
        /// </summary>
        /// <param name="options">Optional polish settings.</param>
        public VisioFluentDocument Polish(VisioDiagramPolishOptions? options = null) {
            _document.PolishDiagrams(options);
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
