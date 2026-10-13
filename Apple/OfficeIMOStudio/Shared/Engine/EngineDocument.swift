import CoreGraphics
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
    /// Whether pages can be rotated, deleted, duplicated, reordered and inserted, and extracted to a new PDF.
    let canMutatePages: Bool?
    let canExtractPages: Bool?
    /// Whether annotations can be added, changed and removed.
    let canEditAnnotations: Bool?
    /// Whether the document permits printing.
    let canPrint: Bool?
    /// Whether pages from other PDFs can be inserted.
    let canImportPages: Bool?
    /// Whether recognized text can be added to the pages.
    let canMakeSearchable: Bool?
    /// Whether the document's form fields can be filled.
    let canFillForms: Bool?
    /// Whether a password protects the document.
    let isEncrypted: Bool?
    /// Whether its password protection can change without the owner password.
    let canChangeProtection: Bool?
    /// Whether a certificate signature can be added.
    let canSign: Bool?

    /// The page image's size at a scale: the page's size times the scale, each side rounded up, as the engine renders it.
    func pageImageSize(page number: Int, scale: Double) -> (width: Int, height: Int) {
        let size = pages[number - 1].displaySize
        return (Int((size.width * scale).rounded(.up)), Int((size.height * scale).rounded(.up)))
    }

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

/// Search hits from officeimo_document_search_json. Rectangles are in points of the page as displayed, from its
/// top-left corner, the space the page layout uses.
struct SearchResult: Decodable, Sendable {
    struct Rectangle: Codable, Sendable, Hashable {
        let x: Double
        let y: Double
        let width: Double
        let height: Double

        var cgRect: CGRect { CGRect(x: x, y: y, width: width, height: height) }
    }

    struct Hit: Decodable, Sendable, Identifiable {
        let page: Int
        let occurrence: Int
        let text: String
        let bounds: Rectangle
        let lineBounds: [Rectangle]

        var id: Int { occurrence }
    }

    let query: String
    let revision: Int64
    let hits: [Hit]
}

/// A copy with a password to open it, or, without userPassword, one without the document's password.
struct ProtectionRequest: Encodable, Sendable {
    let path: String
    var userPassword: String?
    /// Unlocks the permissions below in other apps; the copy then opens with restrictions for everyone else.
    var ownerPassword: String?
    /// The owner password of the open document, when changing its protection needs it.
    var currentOwnerPassword: String?
    var allowPrinting = true
    var allowCopying = true
    var allowChanges = true
}

/// A copy signed with the certificate and private key of a PKCS #12 file.
struct SigningRequest: Encodable, Sendable {
    let path: String
    let certificateURL: URL
    let certificatePassword: String
    var reason: String?
    var location: String?

    enum CodingKeys: String, CodingKey { case path, certificatePath, certificatePassword, reason, location }

    func encode(to encoder: any Encoder) throws {
        var container = encoder.container(keyedBy: CodingKeys.self)
        try container.encode(path, forKey: .path)
        try container.encode(certificateURL.path, forKey: .certificatePath)
        try container.encode(certificatePassword, forKey: .certificatePassword)
        try container.encodeIfPresent(reason, forKey: .reason)
        try container.encodeIfPresent(location, forKey: .location)
    }
}

/// One searched page's hits while a search runs: pagesSearched of pageCount are done.
struct SearchPage: Decodable, Sendable {
    let query: String
    let revision: Int64
    let pagesSearched: Int
    let pageCount: Int
    let hits: [SearchResult.Hit]
}

/// Receives officeimo_document_search_stream_json's pages on the engine's thread.
private final class SearchReceiver: Sendable {
    let pageSearched: @Sendable (SearchPage) -> Void

    init(_ pageSearched: @escaping @Sendable (SearchPage) -> Void) {
        self.pageSearched = pageSearched
    }

    /// The C function the engine calls; its context is a SearchReceiver.
    static let callback: officeimo_search_progress = { context, json in
        guard let context, let json,
              let page = try? JSONDecoder().decode(SearchPage.self, from: Data(String(cString: json).utf8)) else { return }
        Unmanaged<SearchReceiver>.fromOpaque(context).takeUnretainedValue().pageSearched(page)
    }
}

/// A colour with components from 0 through 1, as the engine's JSON carries it.
struct EngineColor: Codable, Sendable, Hashable {
    let red: Double
    let green: Double
    let blue: Double
}

/// A point or rectangle in the displayed page's points, from its top-left corner.
struct EnginePoint: Codable, Sendable, Hashable {
    let x: Double
    let y: Double
}

/// A markup annotation to add, as officeimo_document_add_annotation_json takes it.
struct AnnotationRequest: Encodable, Sendable {
    enum Tool: String, Encodable, Sendable {
        case ink, note, freeText, highlight, underline, strikeOut, squiggly, rectangle, ellipse, line, polygon, link, stamp
    }

    var tool: Tool
    var page: Int
    var bounds: SearchResult.Rectangle?
    var strokes: [[EnginePoint]]?
    var text: String?
    var author: String?
    var color: EngineColor?
    var strokeWidth: Double?
    var opacity: Double?
    /// Text markup's lines, in the displayed page's points.
    var quads: [SearchResult.Rectangle]?
    /// Where a link goes: a web or mail address, or a page of this document.
    var linkUri: String?
    var linkPage: Int?
    /// A standard stamp's name, such as Approved or Draft.
    var stamp: String?
    /// How a line ends: None, OpenArrow or ClosedArrow.
    var lineEnd: String?
}

/// The markup annotations on a page, from officeimo_document_annotations_json.
struct AnnotationList: Decodable, Sendable {
    struct Annotation: Decodable, Sendable, Identifiable, Hashable {
        let objectNumber: Int
        let page: Int
        let subtype: String
        let bounds: SearchResult.Rectangle
        let contents: String?
        let author: String?
        let color: EngineColor?

        var id: Int { objectNumber }
    }

    let revision: Int64
    let page: Int
    let canEdit: Bool
    let annotations: [Annotation]
}

/// A change to an annotation's contents, author or colour, or alone its rectangle on the displayed page, which moves
/// or resizes it; nil fields are kept.
struct AnnotationUpdate: Encodable, Sendable {
    let objectNumber: Int
    var contents: String?
    var author: String?
    var color: EngineColor?
    var bounds: SearchResult.Rectangle?
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
    let handle: officeimo_document
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

    /// Renders a rectangle of the page image at scale pixels per point, for display. The engine's opaque RGBA pixels
    /// become the image's storage without a copy and are released with it. Tiles of one page and scale join exactly.
    func renderRegion(_ number: Int, scale: Double, x: Int, y: Int, width: Int, height: Int) async throws -> CGImage {
        try await withEngineCancellation { cancellation in
            try regionImage(handle, number, scale, x, y, width, height, cancellation)
        }
    }

    /// Renders a whole page from a thread outside the actor, after any call already running. The system's print
    /// renderers draw synchronously, so they wait here rather than awaiting the actor.
    nonisolated func renderPageAndWait(_ number: Int, scale: Double, width: Int, height: Int) throws -> CGImage {
        dispatchPrecondition(condition: .notOnQueue(queue))
        return try queue.sync { try regionImage(handle, number, scale, 0, 0, width, height, 0) }
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

    func deletePages(_ pages: [Int]) async throws {
        try await pageCall(pages) { officeimo_document_delete_pages(handle, $0, $1, $2, $3) }
    }

    func duplicatePages(_ pages: [Int]) async throws {
        try await pageCall(pages) { officeimo_document_duplicate_pages(handle, $0, $1, $2, $3) }
    }

    /// Puts the pages in a new order: order lists every current page number once.
    func reorderPages(_ order: [Int]) async throws {
        try await pageCall(order) { officeimo_document_reorder_pages(handle, $0, $1, $2, $3) }
    }

    /// Inserts a blank page of width by height points before a page, or after the last with the page count + 1.
    func insertBlankPage(before page: Int, width: Double, height: Double) async throws {
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            try check(officeimo_document_insert_blank_page(handle, Int32(page), width, height, cancellation, &error), error)
        }
    }

    /// Writes pages to a new PDF at url; the open document is unchanged.
    func extractPages(_ pages: [Int], to url: URL) async throws {
        let numbers = pages.map(Int32.init)
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            let status = url.path.withCString { path in
                numbers.withUnsafeBufferPointer {
                    officeimo_document_extract_pages(handle, $0.baseAddress, Int32($0.count), path, cancellation, &error)
                }
            }
            try check(status, error)
        }
    }

    /// Recognizes each page's text with Vision and adds it as invisible text, as one undoable change. Progress reports
    /// the page being read and the page count, from the engine's thread.
    func makeSearchable(progress: @escaping @Sendable (_ page: Int, _ pageCount: Int) -> Void) async throws -> SearchableSummary {
        let recognizer = Unmanaged.passRetained(PageRecognizer(progress: progress))
        defer { recognizer.release() }
        return try await withEngineCancellation { cancellation in
            var json: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            let status = officeimo_document_make_searchable_json(handle, PageRecognizer.callback, recognizer.toOpaque(), cancellation, &json, &error)
            let result = takeString(json)
            try check(status, error)
            return try JSONDecoder().decode(SearchableSummary.self, from: Data((result ?? "{}").utf8))
        }
    }

    /// Inserts pages from other PDFs before a page, in the order given, as one undoable change. A source that needs
    /// its password, or was given a wrong one, throws ImportPasswordNeeded naming it, so the app can ask and retry.
    func importPages(_ sources: [ImportSource], before page: Int) async throws -> Int {
        let request = String(decoding: try JSONEncoder().encode(ImportRequest(insertBefore: page, sources: sources)), as: UTF8.self)
        let urls = sources.map(\.url)
        return try await withEngineCancellation { cancellation in
            // Each source stays readable for the call, as a file the user chose may be outside the sandbox.
            let accessing = urls.map { $0.startAccessingSecurityScopedResource() }
            defer { for (url, started) in zip(urls, accessing) where started { url.stopAccessingSecurityScopedResource() } }
            var json: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            let status = request.withCString { officeimo_document_import_pages_json(handle, $0, cancellation, &json, &error) }
            let result = takeString(json)
            if status == EngineStatus.passwordRequired.rawValue || status == EngineStatus.invalidPassword.rawValue,
               let result, let needed = try? JSONDecoder().decode(ImportPasswordNeeded.self, from: Data(result.utf8)) {
                _ = takeString(error)
                var reported = needed
                reported.isRetry = status == EngineStatus.invalidPassword.rawValue
                throw reported
            }
            try check(status, error)
            return try JSONDecoder().decode(ImportResult.self, from: Data((result ?? "{}").utf8)).imported
        }
    }

    /// Finds text, ignoring case.
    func search(_ query: String) async throws -> SearchResult {
        let json: String = try await withEngineCancellation { cancellation in
            var json: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            let status = query.withCString { officeimo_document_search_json(handle, $0, cancellation, &json, &error) }
            try check(status, error)
            return takeString(json) ?? "{}"
        }
        return try JSONDecoder().decode(SearchResult.self, from: Data(json.utf8))
    }

    /// Finds text, ignoring case, passing each page's hits from the engine's thread as soon as the page is searched.
    func search(_ query: String, pageSearched: @escaping @Sendable (SearchPage) -> Void) async throws -> SearchResult {
        let receiver = Unmanaged.passRetained(SearchReceiver(pageSearched))
        defer { receiver.release() }
        let json: String = try await withEngineCancellation { cancellation in
            var json: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            let status = query.withCString {
                officeimo_document_search_stream_json(handle, $0, SearchReceiver.callback, receiver.toOpaque(), cancellation, &json, &error)
            }
            try check(status, error)
            return takeString(json) ?? "{}"
        }
        return try JSONDecoder().decode(SearchResult.self, from: Data(json.utf8))
    }

    func addAnnotation(_ request: AnnotationRequest) async throws {
        let json = String(decoding: try JSONEncoder().encode(request), as: UTF8.self)
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            try check(json.withCString { officeimo_document_add_annotation_json(handle, $0, cancellation, &error) }, error)
        }
    }

    func pageText(page: Int) async throws -> PageText {
        let json: String = try await withEngineCancellation { cancellation in
            var json: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            try check(officeimo_document_page_text_json(handle, Int32(page), cancellation, &json, &error), error)
            return takeString(json) ?? "{}"
        }
        return try JSONDecoder().decode(PageText.self, from: Data(json.utf8))
    }

    func annotations(page: Int) async throws -> AnnotationList {
        let json: String = try await withEngineCancellation { cancellation in
            var json: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            try check(officeimo_document_annotations_json(handle, Int32(page), cancellation, &json, &error), error)
            return takeString(json) ?? "{}"
        }
        return try JSONDecoder().decode(AnnotationList.self, from: Data(json.utf8))
    }

    func updateAnnotation(_ update: AnnotationUpdate) async throws {
        let json = String(decoding: try JSONEncoder().encode(update), as: UTF8.self)
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            try check(json.withCString { officeimo_document_update_annotation_json(handle, $0, cancellation, &error) }, error)
        }
    }

    func removeAnnotation(_ objectNumber: Int) async throws {
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            try check(officeimo_document_remove_annotation(handle, Int32(objectNumber), cancellation, &error), error)
        }
    }

    private func pageCall(_ pages: [Int],
                          _ call: (UnsafePointer<Int32>?, Int32, officeimo_cancellation, UnsafeMutablePointer<UnsafeMutablePointer<CChar>?>) -> Int32) async throws {
        let numbers = pages.map(Int32.init)
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            let status = numbers.withUnsafeBufferPointer { call($0.baseAddress, Int32($0.count), cancellation, &error) }
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

    /// Saves in place from a thread outside the actor, after any call already running. The system's document saving
    /// is synchronous and runs on its own thread, so it waits here rather than awaiting the actor.
    nonisolated func saveInPlaceAndWait() throws {
        dispatchPrecondition(condition: .notOnQueue(queue))
        try queue.sync {
            var error: UnsafeMutablePointer<CChar>?
            try check(officeimo_document_save(handle, nil, 0, &error), error)
        }
    }

    /// Writes the open revision with a password to open it, or without the document's password when protection.userPassword
    /// is nil. A wrong owner password throws EngineError with .invalidPassword.
    func saveProtectedCopy(_ protection: ProtectionRequest) async throws {
        let request = String(decoding: try JSONEncoder().encode(protection), as: UTF8.self)
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            try check(request.withCString { officeimo_document_save_protected_copy_json(handle, $0, cancellation, &error) }, error)
        }
    }

    /// Writes the open revision with a certificate signature. A password the certificate file doesn't accept throws
    /// EngineError with .invalidPassword.
    func saveSignedCopy(_ signing: SigningRequest) async throws {
        let request = String(decoding: try JSONEncoder().encode(signing), as: UTF8.self)
        let certificate = signing.certificateURL
        try await withEngineCancellation { cancellation in
            let accessing = certificate.startAccessingSecurityScopedResource()
            defer { if accessing { certificate.stopAccessingSecurityScopedResource() } }
            var error: UnsafeMutablePointer<CChar>?
            try check(request.withCString { officeimo_document_save_signed_copy_json(handle, $0, cancellation, &error) }, error)
        }
    }

    /// Writes the open revision, with unsaved changes, to url; the document keeps its location and saved state.
    func saveCopy(to url: URL) async throws {
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            try check(url.path.withCString { officeimo_document_save_copy(handle, $0, cancellation, &error) }, error)
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

/// A rendered rectangle of a page as an image over the engine's pixels, which it releases when discarded.
private func regionImage(_ handle: officeimo_document, _ number: Int, _ scale: Double, _ x: Int, _ y: Int, _ width: Int,
                         _ height: Int, _ cancellation: officeimo_cancellation) throws -> CGImage {
    var pixels: UnsafeMutablePointer<UInt8>?
    var length = 0
    var error: UnsafeMutablePointer<CChar>?
    try check(officeimo_document_render_region(handle, Int32(number), scale, Int32(x), Int32(y), Int32(width),
                                                Int32(height), cancellation, &pixels, &length, &error), error)
    guard let pixels else { throw EngineError(status: .failed, code: 1, message: "The engine returned no pixels.") }
    // From here the provider owns the pixels and releases them when the image no longer needs them.
    guard let provider = CGDataProvider(dataInfo: nil, data: pixels, size: length, releaseData: { _, data, _ in
        officeimo_free(UnsafeMutableRawPointer(mutating: data))
    }) else {
        officeimo_free(pixels)
        throw EngineError(status: .failed, code: 1, message: "The page pixels could not be wrapped.")
    }
    guard let image = CGImage(width: width, height: height, bitsPerComponent: 8, bitsPerPixel: 32, bytesPerRow: width * 4,
                              space: CGColorSpace(name: CGColorSpace.sRGB)!,
                              bitmapInfo: CGBitmapInfo(rawValue: CGImageAlphaInfo.noneSkipLast.rawValue),
                              provider: provider, decode: nil, shouldInterpolate: true, intent: .defaultIntent) else {
        throw EngineError(status: .failed, code: 1, message: "The page pixels could not form an image.")
    }
    return image
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

/// A PDF to take pages from: all of them when pages is nil.
struct ImportSource: Encodable, Sendable {
    let url: URL
    var password: String?
    var pages: [Int]?

    enum CodingKeys: String, CodingKey { case path, password, pages }

    func encode(to encoder: any Encoder) throws {
        var container = encoder.container(keyedBy: CodingKeys.self)
        try container.encode(url.path, forKey: .path)
        try container.encodeIfPresent(password, forKey: .password)
        try container.encodeIfPresent(pages, forKey: .pages)
    }
}

private struct ImportRequest: Encodable {
    let insertBefore: Int
    let sources: [ImportSource]
}

private struct ImportResult: Decodable {
    let imported: Int
}

/// A source to import needs its password: its position in the request and its file name.
struct ImportPasswordNeeded: Error, Decodable, Sendable {
    let source: Int
    let name: String
    /// Whether a password was given and was wrong.
    var isRetry = false

    enum CodingKeys: String, CodingKey { case source, name }
}
