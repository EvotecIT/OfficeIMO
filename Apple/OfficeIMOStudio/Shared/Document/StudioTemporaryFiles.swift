import Darwin
import Foundation

/// Temporary document data is private to one process. Returned share/Shortcut outputs may outlive that process,
/// so handed-off outputs and dead-process folders are retained for a day. A second app never removes live documents.
enum StudioTemporaryFiles {
    private static let container = FileManager.default.temporaryDirectory.appending(path: "OfficeIMOStudio", directoryHint: .isDirectory)
    private static let processRoot = container.appending(path: "Process-\(ProcessInfo.processInfo.processIdentifier)-\(UUID().uuidString)", directoryHint: .isDirectory)
    static var root: URL {
        // A long-running process may hand off a new output just before exit. Keep its retention clock current.
        try? FileManager.default.setAttributes([.modificationDate: Date()], ofItemAtPath: processRoot.path)
        return processRoot
    }

    static func removeStale() {
        removeStaleReturnedOutputs()
        guard let folders = try? FileManager.default.contentsOfDirectory(at: container,
            includingPropertiesForKeys: [.isDirectoryKey, .isSymbolicLinkKey, .contentModificationDateKey]) else { return }
        for folder in folders where folder != root {
            let parts = folder.lastPathComponent.split(separator: "-", maxSplits: 2)
            guard parts.count == 3, parts[0] == "Process", let pid = Int32(parts[1]), pid > 0,
                  UUID(uuidString: String(parts[2])) != nil,
                  let values = try? folder.resourceValues(forKeys: [.isDirectoryKey, .isSymbolicLinkKey, .contentModificationDateKey]),
                  values.isDirectory == true, values.isSymbolicLink != true,
                  let modified = values.contentModificationDate, modified < Date().addingTimeInterval(-86_400),
                  kill(pid, 0) == -1, errno == ESRCH else { continue }
            try? FileManager.default.removeItem(at: folder)
        }
    }

    /// A returned file needs time for the system to consume it; input copies and failed actions are removed immediately.
    static func markReturnedOutput(_ url: URL) {
        guard let folder = returnedOutputFolder(url) else { return }
        try? Data().write(to: folder.appending(path: ".returned"), options: .atomic)
    }

    private static func returnedOutputFolder(_ url: URL) -> URL? {
        let folder = url.deletingLastPathComponent()
        guard folder.path.hasPrefix(root.path + "/"), UUID(uuidString: folder.lastPathComponent) != nil,
              ["Scans", "Converted", "Shortcuts"].contains(folder.deletingLastPathComponent().lastPathComponent) else { return nil }
        return folder
    }

    private static func removeStaleReturnedOutputs() {
        for category in ["Scans", "Converted", "Shortcuts"] {
            let parent = root.appending(path: category, directoryHint: .isDirectory)
            guard let folders = try? FileManager.default.contentsOfDirectory(at: parent,
                includingPropertiesForKeys: [.isDirectoryKey, .isSymbolicLinkKey]) else { continue }
            for folder in folders {
                guard UUID(uuidString: folder.lastPathComponent) != nil,
                      let values = try? folder.resourceValues(forKeys: [.isDirectoryKey, .isSymbolicLinkKey]),
                      values.isDirectory == true, values.isSymbolicLink != true,
                      let marker = try? folder.appending(path: ".returned").resourceValues(forKeys: [.contentModificationDateKey]),
                      let returned = marker.contentModificationDate, returned < Date().addingTimeInterval(-86_400) else { continue }
                try? FileManager.default.removeItem(at: folder)
            }
        }
    }

    /// Removes only a generated PDF's UUID staging directory after a caller has imported its contents.
    static func removeGeneratedPDF(_ url: URL) {
        let folder = url.deletingLastPathComponent()
        guard folder.path.hasPrefix(root.path + "/"), UUID(uuidString: folder.lastPathComponent) != nil,
              ["Scans", "Converted"].contains(folder.deletingLastPathComponent().lastPathComponent) else { return }
        try? FileManager.default.removeItem(at: folder)
    }
}
