import Foundation
import OfficeIMOStudioEngine

/// Document facts, decoded from officeimo_document_info_json.
struct DocumentFacts: Decodable, Sendable {
    struct Page: Decodable, Sendable {
        let number: Int
        /// Points before rotation.
        let width: Double
        let height: Double
        let rotation: Int

        /// The page size as displayed, after rotation.
        var displaySize: CGSize {
            rotation % 180 == 0 ? CGSize(width: width, height: height) : CGSize(width: height, height: width)
        }
    }

    let abiVersion: Int
    let fileName: String
    let path: String
    let fileSize: Int64
    let revision: Int64
    let isDirty: Bool
    let canUndo: Bool
    let canRedo: Bool
    let hasRecovery: Bool
    let canExtractText: Bool
    let canExtractContent: Bool
    let pages: [Page]
    /// The largest page image the engine renders; absent from engines that predate it.
    let maxRenderPixels: Int64?

    /// Pixels per point for a page filling width pixels, reduced to keep the image within the engine's budget. The
    /// engine rounds each side up to whole pixels, so the reduced width is checked the same way.
    func renderScale(page number: Int, fillingPixelWidth width: Double) -> Double {
        let size = pages[number - 1].displaySize
        guard let maxRenderPixels, (width * (size.height * width / size.width).rounded(.up)) > Double(maxRenderPixels) else {
            return width / size.width
        }
        var pixelWidth = (Double(maxRenderPixels) * size.width / size.height).squareRoot().rounded(.down)
        while pixelWidth > 1, pixelWidth * (size.height * pixelWidth / size.width).rounded(.up) > Double(maxRenderPixels) {
            pixelWidth -= 1
        }
        // Half a pixel inside, so rounding error cannot carry the width up to the next pixel.
        return (pixelWidth - 0.5) / size.width
    }
}

/// A page rendered by the engine as PNG.
struct RenderedPage: Sendable {
    let png: Data
    let pixelWidth: Int
    let pixelHeight: Int
}

/// One open document. Engine calls block, so the actor runs on its own serial queue rather than the main actor or
/// Swift's cooperative pool; the engine also serializes calls on one document, and separate documents run concurrently.
actor EngineDocument {
    private let handle: officeimo_document
    private let queue: DispatchSerialQueue

    nonisolated var unownedExecutor: UnownedSerialExecutor { queue.asUnownedSerialExecutor() }

    private init(handle: officeimo_document, queue: DispatchSerialQueue) {
        self.handle = handle
        self.queue = queue
    }

    deinit { officeimo_document_close(handle) }

    /// Opens a PDF the app can currently read. The app keeps security-scoped access open for the call.
    static func open(_ url: URL, password: String? = nil) async throws -> EngineDocument {
        let queue = DispatchSerialQueue(label: "OfficeIMO.Document", qos: .userInitiated)
        let handle: officeimo_document = try await withCheckedThrowingContinuation { continuation in
            queue.async {
                continuation.resume(with: Result {
                    try withSecurityScope(url) {
                        var document: officeimo_document = 0
                        var error: UnsafeMutablePointer<CChar>?
                        let status = url.path.withCString { path in
                            withOptionalCString(password) { officeimo_document_open(path, $0, 0, &document, &error) }
                        }
                        try check(status, error)
                        return document
                    }
                })
            }
        }
        return EngineDocument(handle: handle, queue: queue)
    }

    func facts() throws -> DocumentFacts {
        var json: UnsafeMutablePointer<CChar>?
        var error: UnsafeMutablePointer<CChar>?
        try check(officeimo_document_info_json(handle, &json, &error), error)
        return try JSONDecoder().decode(DocumentFacts.self, from: Data((takeString(json) ?? "").utf8))
    }

    /// Renders a page (numbered from 1) at scale pixels per PDF point.
    func renderPage(_ number: Int, scale: Double) async throws -> RenderedPage {
        try await withEngineCancellation { cancellation in
            var png: UnsafeMutablePointer<UInt8>?
            var length = 0
            var width: Int32 = 0
            var height: Int32 = 0
            var error: UnsafeMutablePointer<CChar>?
            try check(officeimo_document_render_page(handle, Int32(number), scale, cancellation, &png, &length, &width, &height, &error), error)
            guard let png else { throw EngineError(status: .failed, code: 1, message: "The engine returned no image.") }
            // The bytes stay owned by the engine until released, so the data releases them when it is discarded.
            let data = Data(bytesNoCopy: png, count: length, deallocator: .custom { pointer, _ in officeimo_free(pointer) })
            return RenderedPage(png: data, pixelWidth: Int(width), pixelHeight: Int(height))
        }
    }

    /// Rotates pages (numbered from 1) by a multiple of 90 degrees as one undoable change.
    func rotatePages(_ pages: [Int], degrees: Int) async throws {
        let numbers = pages.map(Int32.init)
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            let status = numbers.withUnsafeBufferPointer {
                officeimo_document_rotate_pages(handle, $0.baseAddress, Int32($0.count), Int32(degrees), cancellation, &error)
            }
            try check(status, error)
        }
    }

    func undo() async throws {
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            try check(officeimo_document_undo(handle, cancellation, &error), error)
        }
    }

    func redo() async throws {
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            try check(officeimo_document_redo(handle, cancellation, &error), error)
        }
    }

    /// Saves in place when url is nil, otherwise to url, which becomes the document's location. A failed save leaves
    /// the previous file intact.
    func save(to url: URL? = nil) async throws {
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            let status: Int32
            if let url {
                status = try withSecurityScope(url) {
                    url.path.withCString { officeimo_document_save(handle, $0, cancellation, &error) }
                }
            } else {
                status = officeimo_document_save(handle, nil, cancellation, &error)
            }
            try check(status, error)
        }
    }
}

/// Keeps security-scoped access open for the duration of one engine call; the engine never retains it.
private func withSecurityScope<T>(_ url: URL, _ work: () throws -> T) throws -> T {
    let accessing = url.startAccessingSecurityScopedResource()
    defer { if accessing { url.stopAccessingSecurityScopedResource() } }
    return try work()
}

private func withOptionalCString<T>(_ string: String?, _ work: (UnsafePointer<CChar>?) -> T) -> T {
    guard let string else { return work(nil) }
    return string.withCString { work($0) }
}
