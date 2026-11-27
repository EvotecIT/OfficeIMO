namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// The theme new content on this slide is written against: the master's theme with the colour and font schemes
        /// of the slide's or layout's theme override in place.
        /// </summary>
        internal OfficeTheme AuthoringTheme {
            get {
                PowerPointDocument? document = OwnerDocument;
                if (document == null) return OfficeTheme.Default;
                PowerPointThemeScope scope = PowerPointThemeScope.For(PartRef);
                return document.ReadLayoutTheme(scope.Master.IsNone ? document.MasterPartRef(0) : scope.Master, scope.ColorScheme, scope.FontScheme).Theme
                       ?? OfficeTheme.Default;
            }
        }
    }
}
