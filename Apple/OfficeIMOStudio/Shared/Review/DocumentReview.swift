import Foundation
import OfficeIMOStudioEngine

/// A check of the document, as officeimo_document_check_json takes it; exactly one field is set.
struct CheckRequest: Encodable, Sendable {
    var signatures: Bool?
    var standard: String?
    var health: Bool?
}

struct CheckResult: Decodable, Sendable {
    struct Signature: Decodable, Sendable, Identifiable, Hashable {
        let signer: String?
        let field: String?
        let signingTime: String?
        let reason: String?
        let location: String?
        /// valid, intact, invalid, unknown or unchecked.
        let status: String
        let coversWholeDocument: Bool
        let certificate: String?
        let timestamp: String?
        let findings: [String]
        var id: String { (field ?? "") + (signingTime ?? "") + (signer ?? "") }
    }

    struct Signatures: Decodable, Sendable {
        let hasSignatures: Bool
        let signatures: [Signature]
        let findings: [String]
    }

    struct Requirement: Decodable, Sendable, Identifiable, Hashable {
        let id: String
        let name: String
        /// met, missing or unsupported.
        let status: String
        let detail: String
    }

    struct Standard: Decodable, Sendable {
        let name: String
        let isReady: Bool
        let requirements: [Requirement]
    }

    struct Health: Decodable, Sendable {
        let fileSize: Int64
        let pages: Int
        let version: String?
        let canRead: Bool
        let canRewrite: Bool
        let isEncrypted: Bool
        let isSigned: Bool
        let isTagged: Bool
        let hasActiveContent: Bool
        let hasAttachments: Bool
        let recoveredProblems: Int
        let unrecoverableProblems: Int
        let diagnostics: [String]
    }

    let revision: Int64
    let signatures: Signatures?
    let standard: Standard?
    let health: Health?
}

/// The standards a document can be checked against.
enum PDFStandard: String, CaseIterable, Identifiable {
    case pdfa2b, pdfa2u, pdfa2a, pdfa3b, pdfa3u, pdfa3a, pdfa4, pdfua1, pdfua2, pdfx1a, pdfx4

    var id: Self { self }
    var title: String {
        switch self {
        case .pdfa2b: "PDF/A-2b"
        case .pdfa2u: "PDF/A-2u"
        case .pdfa2a: "PDF/A-2a"
        case .pdfa3b: "PDF/A-3b"
        case .pdfa3u: "PDF/A-3u"
        case .pdfa3a: "PDF/A-3a"
        case .pdfa4: "PDF/A-4"
        case .pdfua1: "PDF/UA-1"
        case .pdfua2: "PDF/UA-2"
        case .pdfx1a: "PDF/X-1a"
        case .pdfx4: "PDF/X-4"
        }
    }
}

/// An area to redact, in the displayed page's points, with the text a search found there.
struct RedactionMark: Codable, Sendable, Hashable {
    let page: Int
    let bounds: SearchResult.Rectangle
    var preview: String?
}

/// Finds text to redact, or redacts; exactly one field is set.
struct RedactionRequest: Encodable, Sendable {
    struct Apply: Encodable, Sendable {
        var revision: Int64
        var areas: [RedactionMark]?
        var find: RedactionFind?
        var include: [Int]?
        var removedText: String?
        var sanitize = false
    }
    var find: RedactionFind?
    var apply: Apply?
}

struct RedactionResult: Decodable, Sendable {
    struct Verification: Decodable, Sendable {
        let verified: Bool
        let removed: Int
        let residual: Int
        let inconclusive: Int
        let pages: [Int]
    }
    let revision: Int64
    let marks: [RedactionMark]?
    let verification: Verification?
}

/// The standard stamps every PDF reader draws.
enum StandardStamp: String, CaseIterable, Identifiable {
    case approved = "Approved", draft = "Draft", confidential = "Confidential", final = "Final", forComment = "ForComment",
         notApproved = "NotApproved", experimental = "Experimental", expired = "Expired", topSecret = "TopSecret",
         forPublicRelease = "ForPublicRelease", notForPublicRelease = "NotForPublicRelease", asIs = "AsIs", sold = "Sold",
         departmental = "Departmental"

    var id: Self { self }
    /// Green for approvals and releases, red for the rest, as stamp sets usually draw them.
    var color: EngineColor {
        switch self {
        case .approved, .final, .forPublicRelease: EngineColor(red: 0.1, green: 0.5, blue: 0.2)
        case .draft, .forComment, .experimental, .asIs, .departmental: EngineColor(red: 0.15, green: 0.3, blue: 0.75)
        default: EngineColor(red: 0.8, green: 0.1, blue: 0.1)
        }
    }
    var title: String {
        switch self {
        case .approved: String(localized: "Approved")
        case .draft: String(localized: "Draft")
        case .confidential: String(localized: "Confidential")
        case .final: String(localized: "Final")
        case .forComment: String(localized: "For Comment")
        case .notApproved: String(localized: "Not Approved")
        case .experimental: String(localized: "Experimental")
        case .expired: String(localized: "Expired")
        case .topSecret: String(localized: "Top Secret")
        case .forPublicRelease: String(localized: "For Public Release")
        case .notForPublicRelease: String(localized: "Not for Public Release")
        case .asIs: String(localized: "As Is")
        case .sold: String(localized: "Sold")
        case .departmental: String(localized: "Departmental")
        }
    }
}

extension EngineDocument {
    func checkDocument(_ request: CheckRequest) async throws -> CheckResult {
        let body = String(decoding: try JSONEncoder().encode(request), as: UTF8.self)
        let json: String = try await withEngineCancellation { cancellation in
            var json: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            try check(body.withCString { officeimo_document_check_json(handle, $0, cancellation, &json, &error) }, error)
            return takeString(json) ?? "{}"
        }
        return try JSONDecoder().decode(CheckResult.self, from: Data(json.utf8))
    }

    /// Finds text to redact, or redacts as one verified, undoable change.
    func redact(_ request: RedactionRequest) async throws -> RedactionResult {
        let body = String(decoding: try JSONEncoder().encode(request), as: UTF8.self)
        let json: String = try await withEngineCancellation { cancellation in
            var json: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            try check(body.withCString { officeimo_document_redact_json(handle, $0, cancellation, &json, &error) }, error)
            return takeString(json) ?? "{}"
        }
        return try JSONDecoder().decode(RedactionResult.self, from: Data(json.utf8))
    }
}

/// Redaction the engine refuses because a match can't be removed without touching the text around it.
enum RedactionProblem {
    static func needsWholeBlocks(_ error: any Error) -> Bool {
        (error as? EngineError)?.message.contains("would remove an unselected glyph") == true
    }
}
