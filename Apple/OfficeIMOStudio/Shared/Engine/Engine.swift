import Foundation
import OfficeIMOStudioEngine

/// The status codes of the engine's C interface.
enum EngineStatus: Int32, Sendable {
    case ok = 0
    case failed = 1
    case cancelled = 2
    case invalidArgument = 3
    case passwordRequired = 4
    case invalidPassword = 5
    case unsupported = 6
    case notFound = 7
    case notInitialized = 8
}

/// A failed engine call: its status, so the app can respond without parsing text, and the engine's message.
struct EngineError: LocalizedError, Sendable {
    let status: EngineStatus?
    let code: Int32
    let message: String

    var errorDescription: String? { message }
}

/// Process-wide engine calls. Every call blocks, so each runs on a background queue.
enum Engine {
    private static let queue = DispatchQueue(label: "OfficeIMO.Engine", qos: .userInitiated)

    /// The ABI this app was compiled against.
    static let expectedAbiVersion = Int32(OFFICEIMO_ABI_VERSION)

    /// Reads the library's ABI version; the first call into the library also starts its runtime.
    static func abiVersion() async -> Int32 {
        await run { officeimo_abi_version() }
    }

    static func libraryVersion() async -> String {
        await run { takeString(officeimo_library_version()) ?? "" }
    }

    /// Names the app-owned folder for recovery copies. Call once at launch, before opening documents.
    static func initialize(dataDirectory: URL) async throws {
        try FileManager.default.createDirectory(at: dataDirectory, withIntermediateDirectories: true)
        try await run {
            var error: UnsafeMutablePointer<CChar>?
            let status = dataDirectory.path.withCString { officeimo_initialize($0, &error) }
            try check(status, error)
        }
    }

    /// Makes a PDF at url from scanned page images and the words recognized on them.
    static func createPDFFromScan(_ request: ScanRequest, at url: URL) async throws -> ScanSummary {
        let json = String(decoding: try JSONEncoder().encode(request), as: UTF8.self)
        let path = url.path
        let summary: String = try await run {
            var result: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            let status = json.withCString { request in
                path.withCString { officeimo_create_pdf_from_scan_json(request, $0, 0, &result, &error) }
            }
            try check(status, error)
            return takeString(result) ?? "{}"
        }
        return try JSONDecoder().decode(ScanSummary.self, from: Data(summary.utf8))
    }

    static func run<T: Sendable>(_ work: @escaping @Sendable () throws -> T) async throws -> T {
        try await withCheckedThrowingContinuation { continuation in
            queue.async { continuation.resume(with: Result { try work() }) }
        }
    }

    static func run<T: Sendable>(_ work: @escaping @Sendable () -> T) async -> T {
        await withCheckedContinuation { continuation in
            queue.async { continuation.resume(returning: work()) }
        }
    }
}

/// Turns a status and message from the engine into a thrown error. The message is released either way.
func check(_ status: Int32, _ error: UnsafeMutablePointer<CChar>?) throws {
    let message = takeString(error)
    guard status != EngineStatus.ok.rawValue else { return }
    throw EngineError(status: EngineStatus(rawValue: status), code: status, message: message ?? "The engine reported status \(status).")
}

/// Copies a string the engine returned and releases it.
func takeString(_ pointer: UnsafeMutablePointer<CChar>?) -> String? {
    guard let pointer else { return nil }
    defer { officeimo_free(pointer) }
    return String(cString: pointer)
}

/// Runs a call with an engine cancellation token that follows the current task's cancellation. The engine accepts
/// cancellation from any thread while the call runs.
func withEngineCancellation<T>(
    isolation: isolated (any Actor)? = #isolation,
    _ work: (officeimo_cancellation) throws -> T
) async throws -> T {
    let token = officeimo_cancellation_create()
    defer { officeimo_cancellation_free(token) }
    return try await withTaskCancellationHandler {
        try work(token)
    } onCancel: {
        officeimo_cancellation_cancel(token)
    }
}
