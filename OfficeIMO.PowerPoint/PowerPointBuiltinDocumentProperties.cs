using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Provides access to the built-in package properties for a PowerPoint presentation.
    /// </summary>
    public sealed class PowerPointBuiltinDocumentProperties {
        private readonly PowerPointDocument _document;
        private bool _hasReadOnlyLegacyDateOverrides;
        private DateTime? _readOnlyLegacyCreated;
        private DateTime? _readOnlyLegacyModified;
        private DateTime? _readOnlyLegacyLastPrinted;

        internal PowerPointBuiltinDocumentProperties(PowerPointDocument document) {
            _document = document;
        }

        /// <summary>
        ///     Gets or sets the presentation creator.
        /// </summary>
        public string? Creator {
            get => _document.ReadBuiltinProperty(OfficeXmlNamespaces.DublinCore, "creator");
            set => _document.WriteBuiltinProperty(OfficeXmlNamespaces.DublinCore, "creator", value);
        }

        /// <summary>
        ///     Gets or sets the presentation title.
        /// </summary>
        public string? Title {
            get => _document.ReadBuiltinProperty(OfficeXmlNamespaces.DublinCore, "title");
            set => _document.WriteBuiltinProperty(OfficeXmlNamespaces.DublinCore, "title", value);
        }

        /// <summary>
        ///     Gets or sets the presentation description.
        /// </summary>
        public string? Description {
            get => _document.ReadBuiltinProperty(OfficeXmlNamespaces.DublinCore, "description");
            set => _document.WriteBuiltinProperty(OfficeXmlNamespaces.DublinCore, "description", value);
        }

        /// <summary>
        ///     Gets or sets the presentation category.
        /// </summary>
        public string? Category {
            get => _document.ReadBuiltinProperty(OfficeXmlNamespaces.CoreProperties, "category");
            set => _document.WriteBuiltinProperty(OfficeXmlNamespaces.CoreProperties, "category", value);
        }

        /// <summary>
        ///     Gets or sets keywords used for searching and indexing.
        /// </summary>
        public string? Keywords {
            get => _document.ReadBuiltinProperty(OfficeXmlNamespaces.CoreProperties, "keywords");
            set => _document.WriteBuiltinProperty(OfficeXmlNamespaces.CoreProperties, "keywords", value);
        }

        /// <summary>
        ///     Gets or sets the presentation subject.
        /// </summary>
        public string? Subject {
            get => _document.ReadBuiltinProperty(OfficeXmlNamespaces.DublinCore, "subject");
            set => _document.WriteBuiltinProperty(OfficeXmlNamespaces.DublinCore, "subject", value);
        }

        /// <summary>
        ///     Gets or sets the revision number.
        /// </summary>
        public string? Revision {
            get => _document.ReadBuiltinProperty(OfficeXmlNamespaces.CoreProperties, "revision");
            set => _document.WriteBuiltinProperty(OfficeXmlNamespaces.CoreProperties, "revision", value);
        }

        /// <summary>
        ///     Gets or sets the last user who modified the presentation.
        /// </summary>
        public string? LastModifiedBy {
            get => _document.ReadBuiltinProperty(OfficeXmlNamespaces.CoreProperties, "lastModifiedBy");
            set => _document.WriteBuiltinProperty(OfficeXmlNamespaces.CoreProperties, "lastModifiedBy", value);
        }

        /// <summary>
        ///     Gets or sets the presentation version.
        /// </summary>
        public string? Version {
            get => _document.ReadBuiltinProperty(OfficeXmlNamespaces.CoreProperties, "version");
            set => _document.WriteBuiltinProperty(OfficeXmlNamespaces.CoreProperties, "version", value);
        }

        /// <summary>
        ///     Gets or sets the creation date.
        /// </summary>
        public DateTime? Created {
            get => _hasReadOnlyLegacyDateOverrides
                ? _readOnlyLegacyCreated
                : _document.ReadBuiltinDateProperty(OfficeXmlNamespaces.DublinCoreTerms, "created");
            set => _document.WriteBuiltinDateProperty(OfficeXmlNamespaces.DublinCoreTerms, "created", value);
        }

        /// <summary>
        ///     Gets or sets the last modification date.
        /// </summary>
        public DateTime? Modified {
            get => _hasReadOnlyLegacyDateOverrides
                ? _readOnlyLegacyModified
                : _document.ReadBuiltinDateProperty(OfficeXmlNamespaces.DublinCoreTerms, "modified");
            set => _document.WriteBuiltinDateProperty(OfficeXmlNamespaces.DublinCoreTerms, "modified", value);
        }

        /// <summary>
        ///     Gets or sets the last print date.
        /// </summary>
        public DateTime? LastPrinted {
            get => _hasReadOnlyLegacyDateOverrides
                ? _readOnlyLegacyLastPrinted
                : _document.ReadBuiltinDateProperty(OfficeXmlNamespaces.CoreProperties, "lastPrinted");
            set => _document.WriteBuiltinDateProperty(OfficeXmlNamespaces.CoreProperties, "lastPrinted", value);
        }

        internal void SetReadOnlyLegacyDateOverrides(DateTime? created,
            DateTime? modified, DateTime? lastPrinted) {
            _hasReadOnlyLegacyDateOverrides = true;
            _readOnlyLegacyCreated = AsUtc(created);
            _readOnlyLegacyModified = AsUtc(modified);
            _readOnlyLegacyLastPrinted = AsUtc(lastPrinted);
        }

        private static DateTime? AsUtc(DateTime? value) => value.HasValue
            ? value.Value.ToUniversalTime()
            : null;
    }
}
