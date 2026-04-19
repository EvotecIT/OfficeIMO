using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Lightweight recommendation for a generated deck design alternative.
    /// </summary>
    public sealed class PowerPointDeckDesignRecommendation {
        internal PowerPointDeckDesignRecommendation(PowerPointDeckDesignSummary design,
            int preferenceScore, IReadOnlyList<string> reasons) {
            Design = design;
            PreferenceScore = preferenceScore;
            Reasons = reasons;
        }

        /// <summary>
        ///     Design alternative summary.
        /// </summary>
        public PowerPointDeckDesignSummary Design { get; }

        /// <summary>
        ///     Simple score based on how many explicit brief preferences the alternative satisfies.
        /// </summary>
        public int PreferenceScore { get; }

        /// <summary>
        ///     Short explanations callers can show before choosing or rendering an alternative.
        /// </summary>
        public IReadOnlyList<string> Reasons { get; }

        /// <summary>
        ///     Whether the alternative satisfies at least one explicit brief preference.
        /// </summary>
        public bool MatchesPreferences => PreferenceScore > 0;

        /// <inheritdoc />
        public override string ToString() {
            return Design.Index + ": " + Design.DirectionName + " score " + PreferenceScore;
        }
    }
}
