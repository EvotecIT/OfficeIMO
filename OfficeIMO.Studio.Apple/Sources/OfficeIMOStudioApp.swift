import SwiftUI

@main
struct OfficeIMOStudioApp: App {
    var body: some Scene {
        DocumentGroup(newDocument: {
            guard let url = Bundle.main.url(forResource: "Welcome", withExtension: "pdf"),
                  let bytes = try? Data(contentsOf: url), let document = try? StudioDocument(data: bytes) else {
                preconditionFailure("The bundled welcome document is missing or invalid.")
            }
            return document
        }) { file in
            StudioWorkspace(document: file.document, fileURL: file.fileURL)
        }
        #if os(macOS)
        .defaultSize(width: 1120, height: 800)
        .commands { SidebarCommands() }
        #endif
    }
}
