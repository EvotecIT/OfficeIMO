import Foundation

/// Owns the C buffer contract; the engine has no file-system or UI responsibilities.
enum StudioEngine {
    static let maximumBytes = 64 * 1024 * 1024
    struct EngineError: LocalizedError { let message: String; var errorDescription: String? { message } }
    struct Inspection: Decodable { let pages: Int }

    static func inspect(_ data: Data) throws -> Inspection {
        try JSONDecoder().decode(Inspection.self, from: execute(0, data: data))
    }

    static func addNote(to data: Data, page: Int, text: String) throws -> Data {
        try execute(1, data: data, page: page, note: text)
    }

    private static func execute(_ operation: Int32, data: Data, page: Int = 1, note: String = "") throws -> Data {
        guard data.count <= maximumBytes, let pageNumber = Int32(exactly: page) else {
            throw EngineError(message: "Choose a PDF smaller than 64 MiB.")
        }
        let text = Data(note.utf8)
        guard text.count <= 16384 else { throw EngineError(message: "Enter a shorter note.") }
        var output: UnsafeMutablePointer<UInt8>?
        var count: Int32 = 0
        let status = data.withUnsafeBytes { input in
            text.withUnsafeBytes { noteBuffer in
                oi_studio_pdf(operation, input.bindMemory(to: UInt8.self).baseAddress, Int32(data.count),
                              pageNumber, noteBuffer.bindMemory(to: UInt8.self).baseAddress, Int32(text.count),
                              &output, &count)
            }
        }
        defer { oi_studio_free(output) }
        guard let output, count >= 0, count <= maximumBytes else {
            throw EngineError(message: "The document engine could not complete this operation.")
        }
        let result = Data(bytes: output, count: Int(count))
        guard status == 0 else { throw EngineError(message: String(decoding: result, as: UTF8.self)) }
        return result
    }
}
