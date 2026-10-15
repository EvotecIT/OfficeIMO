import Foundation
import SwiftUI
import UniformTypeIdentifiers

/// An output staged for the system's export panel. The document and any handed-off wrapper retain its files.
struct ExportedFile: Identifiable {
    let document: ExportedDocument
    let contentType: UTType
    let name: String
    let folder: URL
    var notes: String?
    var id: URL { folder }
}

/// A written file or folder. FileWrapper reads its children lazily instead of serializing the whole export into Data.
struct ExportedDocument: FileDocument {
    static let readableContentTypes: [UTType] = [.pdf, .plainText, .markdown, .json, .html, .folder, .data]
    static let writableContentTypes: [UTType] = readableContentTypes + ExportFormat.documents.map(\.contentType)
    private let source: URL
    private let staging: ExportStaging

    init(url: URL, stagingFolder: URL) throws {
        guard url.deletingLastPathComponent().standardizedFileURL == stagingFolder.standardizedFileURL else {
            throw CocoaError(.fileWriteInvalidFileName)
        }
        _ = try url.resourceValues(forKeys: [.isDirectoryKey, .isRegularFileKey])
        source = url
        staging = ExportStaging(folder: stagingFolder)
    }

    init(configuration: ReadConfiguration) throws {
        let folder = StudioTemporaryFiles.root.appending(path: "Exported/\(UUID().uuidString)")
        let url = folder.appending(path: "Document")
        do {
            try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
            try configuration.file.write(to: url, options: .atomic, originalContentsURL: nil)
            try self.init(url: url, stagingFolder: folder)
        } catch {
            try? FileManager.default.removeItem(at: folder)
            throw error
        }
    }

    func fileWrapper(configuration: WriteConfiguration) throws -> FileWrapper {
        // Each write gets its own mutable wrapper; concurrent snapshots share only immutable staged files.
        try StagedExportFileWrapper(url: source, staging: staging)
    }
}

/// Keeps lazy child contents available even if the export panel releases its FileDocument before its FileWrapper.
private final class StagedExportFileWrapper: FileWrapper {
    private let staging: ExportStaging?
    // System consumers decode a standard FileWrapper; the private lease is needed only while reading our staging.
    override var classForCoder: AnyClass { FileWrapper.self }

    init(url: URL, staging: ExportStaging) throws {
        self.staging = staging
        let source = try FileWrapper(url: url, options: [])
        if source.isDirectory, let children = source.fileWrappers {
            // Keep each child's lazy disk contents; never serialize the directory into one aggregate buffer.
            super.init(directoryWithFileWrappers: children)
        } else if source.isRegularFile, let contents = source.regularFileContents {
            super.init(regularFileWithContents: contents)
        } else {
            throw CocoaError(.fileReadCorruptFile)
        }
        if let name = source.preferredFilename { preferredFilename = name }
        fileAttributes = source.fileAttributes
    }

    required init?(coder: NSCoder) {
        staging = nil
        super.init(coder: coder)
    }
}

/// An immutable lease shared by document values and wrappers, released after the last export consumer finishes.
private final class ExportStaging: Sendable {
    let folder: URL
    init(folder: URL) { self.folder = folder }
    deinit { try? FileManager.default.removeItem(at: folder) }
}
