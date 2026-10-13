#if os(iOS)
import SwiftUI
import UIKit

/// The system share sheet for the document. Its Print draws the engine's pages, as the File menu's Print does, in place
/// of the system's Print, which would print the saved file with another renderer. SwiftUI anchors the popover to the
/// toolbar button on iPad and shows a sheet on iPhone.
struct DocumentShareButton: View {
    let session: DocumentSession
    let fileURL: URL
    /// The file being shared: the saved document, or a copy of it with unsaved changes.
    @State private var shared: URL?
    @State private var isPresented = false

    var body: some View {
        Button("Share", systemImage: "square.and.arrow.up") {
            Task {
                do {
                    shared = try await session.shareableFile(saved: fileURL)
                    isPresented = true
                } catch {
                    session.errorMessage = error.localizedDescription
                }
            }
        }
        .popover(isPresented: $isPresented) {
            ActivityView(items: [shared ?? fileURL],
                         activities: session.canPrint ? [PrintActivity { session.print() }] : [],
                         excluded: [.print]) { isPresented = false }
                .presentationCompactAdaptation(.sheet)
                .presentationDetents([.medium, .large])
        }
    }
}

/// UIActivityViewController as the content of a popover or sheet.
private struct ActivityView: UIViewControllerRepresentable {
    let items: [Any]
    let activities: [UIActivity]
    let excluded: [UIActivity.ActivityType]
    let finished: @MainActor () -> Void

    func makeUIViewController(context: Context) -> UIActivityViewController {
        let controller = UIActivityViewController(activityItems: items, applicationActivities: activities)
        controller.excludedActivityTypes = excluded
        let finished = finished
        controller.completionWithItemsHandler = { _, _, _, _ in finished() }
        return controller
    }

    func updateUIViewController(_ controller: UIActivityViewController, context: Context) {}
}
#endif
