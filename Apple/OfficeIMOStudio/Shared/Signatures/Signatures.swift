import CoreGraphics
import Foundation
import Observation

/// A handwritten signature: its strokes scaled so it is one unit tall, from its top-left corner.
struct Signature: Codable, Identifiable, Hashable, Sendable {
    let id: UUID
    let strokes: [[EnginePoint]]

    /// Its width for a height of one.
    var aspect: Double {
        strokes.joined().map(\.x).max() ?? 1
    }

    /// A signature from strokes drawn in any space, scaled to one unit tall from its top-left corner.
    init?(drawn strokes: [[CGPoint]]) {
        let points = strokes.joined()
        guard let minX = points.map(\.x).min(), let maxX = points.map(\.x).max(),
              let minY = points.map(\.y).min(), let maxY = points.map(\.y).max(), maxX - minX > 4 else { return nil }
        let height = max(maxY - minY, 1)
        id = UUID()
        self.strokes = strokes.filter { !$0.isEmpty }.map { stroke in
            stroke.map { EnginePoint(x: Double(($0.x - minX) / height), y: Double(($0.y - minY) / height)) }
        }
    }

    /// The strokes placed on a page: the signature's height in points, centred on a point of the page.
    func strokes(height: Double, centeredAt point: CGPoint) -> [[EnginePoint]] {
        let originX = point.x - aspect * height / 2
        let originY = point.y - height / 2
        return strokes.map { $0.map { EnginePoint(x: originX + $0.x * height, y: originY + $0.y * height) } }
    }
}

/// The signatures kept on this device, as Preview keeps the ones drawn in it.
@MainActor
@Observable
final class SignatureStore {
    static let shared = SignatureStore()

    private(set) var signatures: [Signature] = []
    @ObservationIgnored private let file = FileManager.default.urls(for: .applicationSupportDirectory, in: .userDomainMask)[0]
        .appending(path: "Signatures.json")

    private init() {
        if let data = try? Data(contentsOf: file), let saved = try? JSONDecoder().decode([Signature].self, from: data) {
            signatures = saved
        }
    }

    func add(_ signature: Signature) throws {
        try save(signatures + [signature])
    }

    func remove(_ signature: Signature) throws {
        try save(signatures.filter { $0.id != signature.id })
    }

    private func save(_ next: [Signature]) throws {
        let data = try JSONEncoder().encode(next)
        try FileManager.default.createDirectory(at: file.deletingLastPathComponent(), withIntermediateDirectories: true)
        // Signatures are personal: the file is protected while the device is locked.
        #if os(iOS)
        try data.write(to: file, options: [.atomic, .completeFileProtection])
        #else
        try data.write(to: file, options: .atomic)
        #endif
        signatures = next
    }
}

extension DocumentSession {
    /// Waits for a tap or click on a page to sign there with the signature.
    func beginSigning(with signature: Signature) {
        guard canEditAnnotations else { return }
        if task != .read { task = .read }
        pendingSignature = signature
    }

    /// Signs a page: the signature as ink, centred on a point of the page in its displayed points.
    func placeSignature(page: Int, at point: CGPoint) {
        guard let signature = pendingSignature else { return }
        pendingSignature = nil
        let request = AnnotationRequest(tool: .ink, page: page, strokes: signature.strokes(height: 36, centeredAt: point),
                                        author: Self.author, color: EngineColor(red: 0.05, green: 0.1, blue: 0.35),
                                        strokeWidth: 1.6, opacity: 1)
        addMarkup([request], name: String(localized: "Add Signature"))
    }
}
