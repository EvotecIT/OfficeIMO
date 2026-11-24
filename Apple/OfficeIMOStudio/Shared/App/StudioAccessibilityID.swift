/// Stable accessibility identifiers for the controls the store capture tests drive.
/// They do not change with the interface language, unlike the controls' labels.
/// The capture test bundles compile this file too, so both sides use the same values.
enum StudioAccessibilityID {
    static let pages = "studio.document.pages"
    static let taskPicker = "studio.document.task"
    static let searchButton = "studio.toolbar.search"
    static let searchField = "studio.search.field"
    static let searchClose = "studio.search.close"
    static let searchPosition = "studio.search.position"
    static let sidebarToggle = "studio.toolbar.sidebar"
    static let sidebarMode = "studio.sidebar.mode"
    static let inspectorToggle = "studio.toolbar.inspector"
    static let inspectorFileSize = "studio.inspector.fileSize"
    static let markupDone = "studio.toolbar.markupDone"
    static let viewerPages = "studio.viewer.pages"
}
