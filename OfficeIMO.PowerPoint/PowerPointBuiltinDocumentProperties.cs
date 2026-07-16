using System;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Provides access to the built-in package properties for a PowerPoint presentation.
    /// </summary>
    public sealed class PowerPointBuiltinDocumentProperties {
        private readonly PresentationDocument _presentationDocument;
        private bool _hasReadOnlyLegacyDateOverrides;
        private DateTime? _readOnlyLegacyCreated;
        private DateTime? _readOnlyLegacyModified;
        private DateTime? _readOnlyLegacyLastPrinted;

        internal PowerPointBuiltinDocumentProperties(PresentationDocument presentationDocument) {
            _presentationDocument = presentationDocument ?? throw new ArgumentNullException(nameof(presentationDocument));
        }

        /// <summary>
        ///     Gets or sets the presentation creator.
        /// </summary>
        public string? Creator {
            get => _presentationDocument.PackageProperties.Creator;
            set => _presentationDocument.PackageProperties.Creator = value;
        }

        /// <summary>
        ///     Gets or sets the presentation title.
        /// </summary>
        public string? Title {
            get => _presentationDocument.PackageProperties.Title;
            set => _presentationDocument.PackageProperties.Title = value;
        }

        /// <summary>
        ///     Gets or sets the presentation description.
        /// </summary>
        public string? Description {
            get => _presentationDocument.PackageProperties.Description;
            set => _presentationDocument.PackageProperties.Description = value;
        }

        /// <summary>
        ///     Gets or sets the presentation category.
        /// </summary>
        public string? Category {
            get => _presentationDocument.PackageProperties.Category;
            set => _presentationDocument.PackageProperties.Category = value;
        }

        /// <summary>
        ///     Gets or sets keywords used for searching and indexing.
        /// </summary>
        public string? Keywords {
            get => _presentationDocument.PackageProperties.Keywords;
            set => _presentationDocument.PackageProperties.Keywords = value;
        }

        /// <summary>
        ///     Gets or sets the presentation subject.
        /// </summary>
        public string? Subject {
            get => _presentationDocument.PackageProperties.Subject;
            set => _presentationDocument.PackageProperties.Subject = value;
        }

        /// <summary>
        ///     Gets or sets the revision number.
        /// </summary>
        public string? Revision {
            get => _presentationDocument.PackageProperties.Revision;
            set => _presentationDocument.PackageProperties.Revision = value;
        }

        /// <summary>
        ///     Gets or sets the last user who modified the presentation.
        /// </summary>
        public string? LastModifiedBy {
            get => _presentationDocument.PackageProperties.LastModifiedBy;
            set => _presentationDocument.PackageProperties.LastModifiedBy = value;
        }

        /// <summary>
        ///     Gets or sets the presentation version.
        /// </summary>
        public string? Version {
            get => _presentationDocument.PackageProperties.Version;
            set => _presentationDocument.PackageProperties.Version = value;
        }

        /// <summary>
        ///     Gets or sets the creation date.
        /// </summary>
        public DateTime? Created {
            get => _hasReadOnlyLegacyDateOverrides
                ? _readOnlyLegacyCreated
                : _presentationDocument.PackageProperties.Created;
            set => _presentationDocument.PackageProperties.Created = value;
        }

        /// <summary>
        ///     Gets or sets the last modification date.
        /// </summary>
        public DateTime? Modified {
            get => _hasReadOnlyLegacyDateOverrides
                ? _readOnlyLegacyModified
                : _presentationDocument.PackageProperties.Modified;
            set => _presentationDocument.PackageProperties.Modified = value;
        }

        /// <summary>
        ///     Gets or sets the last print date.
        /// </summary>
        public DateTime? LastPrinted {
            get => _hasReadOnlyLegacyDateOverrides
                ? _readOnlyLegacyLastPrinted
                : _presentationDocument.PackageProperties.LastPrinted;
            set => _presentationDocument.PackageProperties.LastPrinted = value;
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
