#if os(iOS)
import SwiftUI
import UIKit

/// The one-time first screen on iPad and iPhone: the main jobs, each of which starts straight away.
struct WelcomeView: View {
    private static let shownKey = "WelcomeShown"

    let start: (@MainActor @escaping () -> Void) -> Void

    var body: some View {
        NavigationStack {
            List {
                Section {
                    job("Convert to PDF", "arrow.right.doc.on.clipboard", "Word, Excel, PowerPoint and more") {
                        LaunchDocuments.open { try await ConversionSources.newDocumentFromFile() }
                    }
                    job("Combine Files", "doc.on.doc", "PDFs, images and documents into one PDF") {
                        LaunchDocuments.open { try await ConversionSources.newDocumentFromFiles() }
                    }
                    if ScanSources.isCameraAvailable {
                        job("Scan Document", "doc.viewfinder", "Paper into a searchable PDF") {
                            LaunchDocuments.open { try await ScanSources.newDocumentFromCamera() }
                        }
                    }
                    job("Fill & Sign", "signature", "Try it on a sample form") {
                        LaunchDocuments.open { SampleDocument.url }
                    }
                } header: {
                    Text("Read, edit, convert and sign documents on your device.")
                        .textCase(nil)
                } footer: {
                    Text("Everything happens on your device. Nothing is uploaded and no account is needed.")
                }
            }
            .navigationTitle("Welcome")
            .toolbar {
                ToolbarItem(placement: .confirmationAction) {
                    Button("Done") { start {} }
                }
            }
        }
    }

    private func job(_ title: LocalizedStringKey, _ symbol: String, _ detail: LocalizedStringKey,
                     action: @MainActor @escaping () -> Void) -> some View {
        Button { start(action) } label: {
            Label {
                VStack(alignment: .leading, spacing: 1) {
                    Text(title)
                    Text(detail).font(.footnote).foregroundStyle(.secondary)
                }
            } icon: {
                Image(systemName: symbol)
            }
        }
        .foregroundStyle(.primary)
    }

    /// Shows the screen on the first launch, and not to unattended runs.
    @MainActor
    static func presentIfNeeded() {
        let defaults = UserDefaults.standard
        guard !defaults.bool(forKey: shownKey), !Autorun.isRequested,
              !ProcessInfo.processInfo.arguments.contains("-OfficeIMOSkipWelcome") else { return }
        defaults.set(true, forKey: shownKey)
        var host: UIHostingController<WelcomeView>?
        host = UIHostingController(rootView: WelcomeView { action in
            host?.dismiss(animated: true) { action() }
            host = nil
        })
        Task { if let host { _ = await ScanSources.present(host) } }
    }
}
#endif
