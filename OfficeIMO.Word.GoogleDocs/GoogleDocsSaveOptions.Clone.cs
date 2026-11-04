using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    public sealed partial class GoogleDocsSaveOptions {
        /// <summary>Creates an independent copy of the mutable Google Docs export settings.</summary>
        public GoogleDocsSaveOptions Clone() {
            if (Location == null || FidelityPolicy == null || UnsupportedFeatures == null || Tabs == null ||
                Replace == null || RasterFallbackImageOptions == null)
                throw new ArgumentException("Google Docs export option groups cannot be null.");
            var copy = new GoogleDocsSaveOptions {
                Title = Title,
                Location = new GoogleDriveFileLocation {
                    DriveId = Location.DriveId, FolderId = Location.FolderId,
                    ExistingFileId = Location.ExistingFileId, SharedDriveAware = Location.SharedDriveAware
                },
                FidelityPolicy = new GoogleWorkspaceFidelityPolicy { PreflightMode = FidelityPolicy.PreflightMode },
                UnsupportedFeatures = new GoogleDocsUnsupportedFeatureOptions {
                    FloatingContent = UnsupportedFeatures.FloatingContent, Charts = UnsupportedFeatures.Charts,
                    SmartArt = UnsupportedFeatures.SmartArt, ContentControls = UnsupportedFeatures.ContentControls,
                    EmbeddedObjects = UnsupportedFeatures.EmbeddedObjects, Watermarks = UnsupportedFeatures.Watermarks,
                    Comments = UnsupportedFeatures.Comments, Equations = UnsupportedFeatures.Equations
                },
                InlineImageMode = InlineImageMode,
                Tabs = new GoogleDocsTabOptions { Strategy = Tabs.Strategy, TabId = Tabs.TabId },
                Replace = new GoogleDocsReplaceOptions {
                    ConflictMode = Replace.ConflictMode, ExpectedRevisionId = Replace.ExpectedRevisionId
                },
                Comments = Comments,
                RasterFallbackImageOptions = RasterFallbackImageOptions.Clone()
            };
            foreach (string code in FidelityPolicy.AcceptedDiagnosticCodes) copy.FidelityPolicy.AcceptedDiagnosticCodes.Add(code);
            return copy;
        }
    }
}
