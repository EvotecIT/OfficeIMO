namespace OfficeIMO.Word.Html {
    /// <summary>Controls which tracked-change view is projected into Word review HTML.</summary>
    public enum WordTrackedChangeExportPolicy {
        /// <summary>Projects inserted and moved-to content and omits deleted and moved-from content.</summary>
        Final,

        /// <summary>Projects deleted and moved-from content and omits inserted and moved-to content.</summary>
        Original,

        /// <summary>Projects both views as inert HTML <c>ins</c>/<c>del</c> markup with review metadata.</summary>
        Markup
    }

    /// <summary>Controls how live Word fields are represented in static review HTML.</summary>
    public enum WordFieldExportPolicy {
        /// <summary>Exports only the current stored field result.</summary>
        VisibleResult,

        /// <summary>Exports the stored result and an inert inventory of field instructions and state.</summary>
        VisibleResultWithReviewMetadata
    }
}
