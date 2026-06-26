namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents worksheet sort dialog metadata parsed from a legacy BIFF Sort record.
    /// </summary>
    public sealed class LegacyXlsSortSettings {
        /// <summary>
        /// Creates sort settings parsed from a BIFF Sort record.
        /// </summary>
        /// <param name="sortLeftToRight">Whether the sort operates from left to right instead of top to bottom.</param>
        /// <param name="key1Descending">Whether the first sort key is descending.</param>
        /// <param name="key2Descending">Whether the second sort key is descending.</param>
        /// <param name="key3Descending">Whether the third sort key is descending.</param>
        /// <param name="caseSensitive">Whether sorting is case-sensitive.</param>
        /// <param name="customListIndex">Zero-based custom-list sort order index.</param>
        /// <param name="usePhoneticInformation">Whether phonetic information participates in sorting.</param>
        /// <param name="key1">First sort key label, when present.</param>
        /// <param name="key2">Second sort key label, when present.</param>
        /// <param name="key3">Third sort key label, when present.</param>
        public LegacyXlsSortSettings(
            bool sortLeftToRight,
            bool key1Descending,
            bool key2Descending,
            bool key3Descending,
            bool caseSensitive,
            int customListIndex,
            bool usePhoneticInformation,
            string? key1,
            string? key2,
            string? key3) {
            SortLeftToRight = sortLeftToRight;
            Key1Descending = key1Descending;
            Key2Descending = key2Descending;
            Key3Descending = key3Descending;
            CaseSensitive = caseSensitive;
            CustomListIndex = customListIndex;
            UsePhoneticInformation = usePhoneticInformation;
            Key1 = key1;
            Key2 = key2;
            Key3 = key3;
        }

        /// <summary>Gets whether the sort operates from left to right instead of top to bottom.</summary>
        public bool SortLeftToRight { get; }

        /// <summary>Gets whether the first sort key is descending.</summary>
        public bool Key1Descending { get; }

        /// <summary>Gets whether the second sort key is descending.</summary>
        public bool Key2Descending { get; }

        /// <summary>Gets whether the third sort key is descending.</summary>
        public bool Key3Descending { get; }

        /// <summary>Gets whether sorting is case-sensitive.</summary>
        public bool CaseSensitive { get; }

        /// <summary>Gets the zero-based custom-list sort order index.</summary>
        public int CustomListIndex { get; }

        /// <summary>Gets whether phonetic information participates in sorting.</summary>
        public bool UsePhoneticInformation { get; }

        /// <summary>Gets the first sort key label, when present.</summary>
        public string? Key1 { get; }

        /// <summary>Gets the second sort key label, when present.</summary>
        public string? Key2 { get; }

        /// <summary>Gets the third sort key label, when present.</summary>
        public string? Key3 { get; }
    }
}
