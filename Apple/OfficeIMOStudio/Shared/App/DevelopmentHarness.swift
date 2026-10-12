#if DEBUG && os(iOS)
import SwiftUI

/// Development builds only. `-OfficeIMOOpen <file>` opens Documents/<file> from the app's container straight into a
/// document window. The simulator's document browser cannot read or create files when the simulator's data lives on
/// an external volume, so this exercises the window without it; the system's document handling is unchanged.
enum DevelopmentHarness {
    static var file: String? {
        let arguments = ProcessInfo.processInfo.arguments
        guard let index = arguments.firstIndex(of: "-OfficeIMOOpen"), index + 1 < arguments.count else { return nil }
        return arguments[index + 1]
    }
}

struct DevelopmentHarnessApp: App {
    private let url: URL
    private let document: StudioDocument?

    init() {
        StudioDocument.removeStaleWorkingCopies()
        let documents = FileManager.default.urls(for: .documentDirectory, in: .userDomainMask)[0]
        url = documents.appending(path: DevelopmentHarness.file ?? "")
        document = try? StudioDocument(contentsOf: url)
    }

    var body: some Scene {
        WindowGroup {
            if let document {
                HarnessWindow(document: document, url: url)
            } else {
                ContentUnavailableView("No Document", systemImage: "doc", description: Text(url.lastPathComponent))
            }
        }
        .commands { StudioCommands() }
    }
}

/// The navigation bar the system's document group adds in compact width; regular width uses the split view's bars.
private struct HarnessWindow: View {
    let document: StudioDocument
    let url: URL
    @Environment(\.horizontalSizeClass) private var sizeClass

    var body: some View {
        if sizeClass == .compact {
            NavigationStack {
                window.navigationBarTitleDisplayMode(.inline)
            }
        } else {
            window
        }
    }

    private var window: some View {
        DocumentWindow(document: document, fileURL: url, title: url.deletingPathExtension().lastPathComponent)
    }
}
#endif
