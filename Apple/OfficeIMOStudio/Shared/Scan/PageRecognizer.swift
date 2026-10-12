import Foundation
import OfficeIMOStudioEngine
import Synchronization

/// What making a document searchable added.
struct SearchableSummary: Decodable, Sendable {
    /// Pages that gained text.
    let pages: Int
    let words: Int
}

/// The engine's text recognizer for officeimo_document_make_searchable_json: the engine renders each page and asks
/// Vision, through this, for the words on it. The engine calls it on its own thread and waits, so it waits in turn
/// for the same recognition the scanner uses.
final class PageRecognizer: Sendable {
    private let progress: @Sendable (_ page: Int, _ pageCount: Int) -> Void

    init(progress: @escaping @Sendable (_ page: Int, _ pageCount: Int) -> Void) {
        self.progress = progress
    }

    /// The C function the engine calls; its context is a PageRecognizer.
    static let callback: officeimo_text_recognizer = { context, page, pageCount, png, length, wordsJson in
        guard let context, let png, let wordsJson else { return EngineStatus.failed.rawValue }
        let recognizer = Unmanaged<PageRecognizer>.fromOpaque(context).takeUnretainedValue()
        return recognizer.recognize(page: Int(page), of: Int(pageCount), png: Data(bytes: png, count: length), into: wordsJson)
    }

    private func recognize(page: Int, of pageCount: Int, png: Data, into wordsJson: UnsafeMutablePointer<UnsafeMutablePointer<CChar>?>) -> Int32 {
        progress(page, pageCount)
        guard let image = ScanBuilder.image(from: png) else { return EngineStatus.failed.rawValue }
        let words = Self.wait { await ScanBuilder.recognizeWords(in: image) }
        guard let data = try? JSONEncoder().encode(words) else { return EngineStatus.failed.rawValue }
        // The engine frees the words with free(), so they are copied into memory from strdup.
        wordsJson.pointee = strdup(String(decoding: data, as: UTF8.self))
        return wordsJson.pointee == nil ? EngineStatus.failed.rawValue : EngineStatus.ok.rawValue
    }

    /// Runs async work to completion from a thread that is not Swift's cooperative pool, as the engine's thread is.
    private static func wait<T: Sendable>(_ work: @escaping @Sendable () async -> T) -> T {
        let result = Mutex<T?>(nil)
        let done = DispatchSemaphore(value: 0)
        Task.detached {
            let value = await work()
            result.withLock { $0 = value }
            done.signal()
        }
        done.wait()
        return result.withLock { $0! }
    }
}
