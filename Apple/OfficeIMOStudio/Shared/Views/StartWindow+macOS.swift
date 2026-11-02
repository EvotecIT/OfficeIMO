#if os(macOS)
import AppKit
import SwiftUI

/// The first window on the Mac: what the app does, and recent documents. The document scene's own launch is an open
/// panel, which says nothing about converting or combining.
struct StartWindow: View {
    static let id = "start"

    @Environment(\.newDocument) private var newDocument
    private let recents = NSDocumentController.shared.recentDocumentURLs

    var body: some View {
        HStack(spacing: 0) {
            VStack(alignment: .leading, spacing: 4) {
                Image(nsImage: NSApp.applicationIconImage)
                    .resizable()
                    .frame(width: 64, height: 64)
                    .accessibilityHidden(true)
                Text(verbatim: "OfficeIMO Studio")
                    .font(.title.weight(.semibold))
                Text("Read, edit, convert and sign documents on your Mac.")
                    .foregroundStyle(.secondary)
                    .padding(.bottom, 16)
                action("Open a PDF…", "doc.viewfinder") { NSDocumentController.shared.openDocument(nil) }
                action("Convert to PDF…", "arrow.right.doc.on.clipboard", detail: "Word, Excel, PowerPoint and more") {
                    ConversionSources.newDocumentFromFile(newDocument)
                }
                action("Combine Files…", "doc.on.doc", detail: "PDFs, images and documents into one PDF") {
                    ConversionSources.newDocumentFromFiles(newDocument)
                }
                action("New PDF from Images…", "photo") { ScanSources.newDocumentFromImageFiles(newDocument) }
                action("New Blank PDF", "doc") { newDocument { StudioDocument() } }
                action("Try a Sample PDF", "sparkles") { SampleDocument.open(newDocument) }
                Spacer(minLength: 0)
            }
            .padding(24)
            .frame(width: 380, alignment: .leading)
            Divider()
            VStack(alignment: .leading, spacing: 8) {
                Text("Recent")
                    .font(.headline)
                if recents.isEmpty {
                    Spacer()
                    Text("No recent documents")
                        .foregroundStyle(.secondary)
                        .frame(maxWidth: .infinity)
                    Spacer()
                } else {
                    List(recents, id: \.self) { url in
                        Button { open(url) } label: {
                            Label(url.deletingPathExtension().lastPathComponent, systemImage: "doc.richtext")
                                .lineLimit(1)
                        }
                        .buttonStyle(.plain)
                    }
                    .listStyle(.plain)
                }
            }
            .padding(24)
            .frame(width: 280)
        }
        .frame(height: 460)
    }

    private func action(_ title: LocalizedStringKey, _ symbol: String, detail: LocalizedStringKey? = nil,
                        perform: @escaping () -> Void) -> some View {
        Button(action: perform) {
            Label {
                VStack(alignment: .leading, spacing: 1) {
                    Text(title)
                    if let detail {
                        Text(detail).font(.caption).foregroundStyle(.secondary)
                    }
                }
            } icon: {
                Image(systemName: symbol).frame(width: 24)
            }
            .frame(maxWidth: .infinity, alignment: .leading)
            .padding(.vertical, 6)
            .contentShape(.rect)
        }
        .buttonStyle(.plain)
    }

    private func open(_ url: URL) {
        NSDocumentController.shared.openDocument(withContentsOf: url, display: true) { _, _, error in
            if let error { NSAlert(error: error).runModal() }
        }
    }
}

/// Keeps the Mac from opening a panel or an untitled document at launch, as the start window takes their place.
final class StudioAppDelegate: NSObject, NSApplicationDelegate {
    func applicationShouldOpenUntitledFile(_ sender: NSApplication) -> Bool { false }
}
#endif
