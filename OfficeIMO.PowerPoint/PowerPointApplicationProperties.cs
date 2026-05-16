using System;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Provides access to extended application properties for a PowerPoint presentation.
    /// </summary>
    public sealed class PowerPointApplicationProperties {
        private readonly PresentationDocument _presentationDocument;

        internal PowerPointApplicationProperties(PresentationDocument presentationDocument) {
            _presentationDocument = presentationDocument ?? throw new ArgumentNullException(nameof(presentationDocument));
        }

        /// <summary>
        ///     Gets or sets the application name that created the presentation.
        /// </summary>
        public string Application {
            get => GetProperties().Application?.Text ?? string.Empty;
            set {
                Ap.Properties properties = GetProperties();
                properties.Application ??= new Application();
                properties.Application.Text = value;
            }
        }

        /// <summary>
        ///     Gets or sets the application version that created the presentation.
        /// </summary>
        public string ApplicationVersion {
            get => GetProperties().ApplicationVersion?.Text ?? string.Empty;
            set {
                Ap.Properties properties = GetProperties();
                properties.ApplicationVersion ??= new ApplicationVersion();
                properties.ApplicationVersion.Text = value;
            }
        }

        /// <summary>
        ///     Gets or sets the company associated with the presentation.
        /// </summary>
        public string Company {
            get => GetProperties().Company?.Text ?? string.Empty;
            set {
                Ap.Properties properties = GetProperties();
                properties.Company ??= new Company();
                properties.Company.Text = value;
            }
        }

        /// <summary>
        ///     Gets or sets the manager associated with the presentation.
        /// </summary>
        public string Manager {
            get => GetProperties().Manager?.Text ?? string.Empty;
            set {
                Ap.Properties properties = GetProperties();
                properties.Manager ??= new Manager();
                properties.Manager.Text = value;
            }
        }

        /// <summary>
        ///     Gets or sets the presentation format description.
        /// </summary>
        public string PresentationFormat {
            get => GetProperties().PresentationFormat?.Text ?? string.Empty;
            set {
                Ap.Properties properties = GetProperties();
                properties.PresentationFormat ??= new PresentationFormat();
                properties.PresentationFormat.Text = value;
            }
        }

        /// <summary>
        ///     Gets the stored slide count.
        /// </summary>
        public string Slides => GetProperties().Slides?.Text ?? string.Empty;

        /// <summary>
        ///     Gets the stored notes count.
        /// </summary>
        public string Notes => GetProperties().Notes?.Text ?? string.Empty;

        /// <summary>
        ///     Gets the stored hidden-slide count.
        /// </summary>
        public string HiddenSlides => GetProperties().HiddenSlides?.Text ?? string.Empty;

        /// <summary>
        ///     Gets or sets the stored total editing time.
        /// </summary>
        public string TotalTime {
            get => GetProperties().TotalTime?.Text ?? string.Empty;
            set {
                Ap.Properties properties = GetProperties();
                properties.TotalTime ??= new TotalTime();
                properties.TotalTime.Text = value;
            }
        }

        private Ap.Properties GetProperties() {
            ExtendedFilePropertiesPart part = _presentationDocument.ExtendedFilePropertiesPart
                ?? _presentationDocument.AddExtendedFilePropertiesPart();
            part.Properties ??= new Ap.Properties();
            return part.Properties;
        }
    }
}
