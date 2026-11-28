namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Provides access to extended application properties for a PowerPoint presentation.
    /// </summary>
    public sealed class PowerPointApplicationProperties {
        private readonly PowerPointDocument _document;

        internal PowerPointApplicationProperties(PowerPointDocument document) {
            _document = document;
        }

        /// <summary>
        ///     Gets or sets the application name that created the presentation.
        /// </summary>
        public string Application {
            get => _document.ReadApplicationProperty("Application");
            set => _document.WriteApplicationProperty("Application", value);
        }

        /// <summary>
        ///     Gets or sets the application version that created the presentation.
        /// </summary>
        public string ApplicationVersion {
            get => _document.ReadApplicationProperty("AppVersion");
            set => _document.WriteApplicationProperty("AppVersion", value);
        }

        /// <summary>
        ///     Gets or sets the company associated with the presentation.
        /// </summary>
        public string Company {
            get => _document.ReadApplicationProperty("Company");
            set => _document.WriteApplicationProperty("Company", value);
        }

        /// <summary>
        ///     Gets or sets the manager associated with the presentation.
        /// </summary>
        public string Manager {
            get => _document.ReadApplicationProperty("Manager");
            set => _document.WriteApplicationProperty("Manager", value);
        }

        /// <summary>
        ///     Gets or sets the presentation format description.
        /// </summary>
        public string PresentationFormat {
            get => _document.ReadApplicationProperty("PresentationFormat");
            set => _document.WriteApplicationProperty("PresentationFormat", value);
        }

        /// <summary>
        ///     Gets the stored slide count.
        /// </summary>
        public string Slides => _document.ReadApplicationProperty("Slides");

        /// <summary>
        ///     Gets the stored notes count.
        /// </summary>
        public string Notes => _document.ReadApplicationProperty("Notes");

        /// <summary>
        ///     Gets the stored hidden-slide count.
        /// </summary>
        public string HiddenSlides => _document.ReadApplicationProperty("HiddenSlides");

        /// <summary>
        ///     Gets or sets the stored total editing time.
        /// </summary>
        public string TotalTime {
            get => _document.ReadApplicationProperty("TotalTime");
            set => _document.WriteApplicationProperty("TotalTime", value);
        }
    }
}
