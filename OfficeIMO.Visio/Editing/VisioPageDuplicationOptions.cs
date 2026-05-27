namespace OfficeIMO.Visio {
    /// <summary>
    /// Options controlling how a Visio page is duplicated.
    /// </summary>
    public sealed class VisioPageDuplicationOptions {
        /// <summary>
        /// Optional name for the duplicated foreground or background page. When omitted, OfficeIMO creates a unique copy name.
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// When true, a foreground page that uses a background page receives an independent copy of that background page.
        /// When false, the duplicated foreground page keeps using the original background page.
        /// </summary>
        public bool DuplicateBackgroundPage { get; set; }

        /// <summary>
        /// Optional name for the duplicated background page when <see cref="DuplicateBackgroundPage"/> is true.
        /// </summary>
        public string? BackgroundPageName { get; set; }
    }
}
