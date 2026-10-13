import Foundation
import OfficeIMOStudioEngine

/// The document's properties, outline and attachments from the engine, with what the document lets the app change.
struct DocumentStructure: Decodable, Sendable {
    struct Properties: Codable, Sendable, Equatable {
        var title: String?
        var author: String?
        var subject: String?
        var keywords: String?
        var created: String?
        var modified: String?
    }

    /// An outline entry as the engine lists it. Entries without an id have no page and can't be changed.
    struct Entry: Decodable, Sendable {
        let id: String?
        let title: String
        let page: Int?
        let children: [Entry]
    }

    struct Attachment: Decodable, Sendable, Identifiable, Hashable {
        let fileName: String
        let size: Int64
        let description: String?
        let mimeType: String?
        var id: String { fileName }
    }

    let revision: Int64
    let properties: Properties
    let canEditProperties: Bool
    let outline: [Entry]
    let canEditOutline: Bool
    let attachments: [Attachment]
    let canEditAttachments: Bool
    let canCrop: Bool
    let canResize: Bool
    let canStampPages: Bool
    let hasForms: Bool
    let canFlattenForms: Bool

    /// The outline as nodes a list can show, each with a key that stays unique for entries without an id.
    var outlineNodes: [OutlineNode] { OutlineNode.nodes(outline, path: "") }
}

/// An outline entry for the sidebar's disclosure list.
struct OutlineNode: Identifiable, Hashable, Sendable {
    let id: String
    let bookmark: String?
    let title: String
    let page: Int?
    let children: [OutlineNode]?

    static func nodes(_ entries: [DocumentStructure.Entry], path: String) -> [OutlineNode] {
        entries.enumerated().map { index, entry in
            let key = entry.id ?? "\(path)/\(index)"
            let children = nodes(entry.children, path: key)
            return OutlineNode(id: key, bookmark: entry.id, title: entry.title, page: entry.page, children: children.isEmpty ? nil : children)
        }
    }

    /// The top-level entries with pages, in page order, as the parts a split at each bookmark makes.
    static func parts(_ nodes: [OutlineNode], pageCount: Int) -> [SplitPart] {
        let starts = nodes.compactMap { node in node.page.map { (page: $0, title: node.title) } }
            .sorted { $0.page < $1.page }
        var parts: [SplitPart] = []
        if let first = starts.first, first.page > 1 { parts.append(SplitPart(pages: Array(1..<first.page), name: nil)) }
        for (index, start) in starts.enumerated() {
            let end = index + 1 < starts.count ? starts[index + 1].page - 1 : pageCount
            if end >= start.page { parts.append(SplitPart(pages: Array(start.page...end), name: start.title)) }
        }
        return parts
    }
}

/// One edit, as one undoable change; exactly one field is set.
struct EditRequest: Encodable, Sendable {
    var properties: DocumentStructure.Properties?
    var outline: [OutlineEdit]?
    var crop: CropEdit?
    var pageSize: PageSizeEdit?
    var headerFooter: HeaderFooterEdit?
    var watermark: WatermarkEdit?
    var bates: BatesEdit?
    var flatten: FlattenEdit?
    var addAttachment: AttachmentAdd?
    var removeAttachment: String?
}

struct OutlineEdit: Encodable, Sendable {
    enum Op: String, Encodable, Sendable { case add, rename, remove, move, retarget, rebuild }
    var op: Op
    var id: String?
    var title: String?
    var page: Int?
    var parentId: String?
    var index: Int?
}

/// Margins in points cut from each edge of the pages as displayed.
struct CropEdit: Encodable, Sendable {
    var pages: [Int]
    var left: Double
    var top: Double
    var right: Double
    var bottom: Double
}

struct PageSizeEdit: Encodable, Sendable {
    enum Mode: String, Encodable, Sendable, CaseIterable, Identifiable {
        case fit, fill, stretch
        var id: Self { self }
    }
    var pages: [Int]
    var width: Double
    var height: Double
    var mode: Mode
}

/// Text accepts {page}, {pages}, {date}, {file} and {title}.
struct HeaderFooterEdit: Encodable, Sendable {
    var headerLeft: String?
    var headerCenter: String?
    var headerRight: String?
    var footerLeft: String?
    var footerCenter: String?
    var footerRight: String?
    var fontSize: Double?
    var startNumber: Int?
    var pages: [Int]?
}

struct WatermarkEdit: Encodable, Sendable {
    var text: String
    var fontSize: Double?
    var opacity: Double?
    var rotation: Double?
    var behindContent: Bool
    var color: EngineColor?
    var pages: [Int]?
}

struct BatesEdit: Encodable, Sendable {
    enum Position: String, Encodable, Sendable, CaseIterable, Identifiable {
        case topLeft, topCenter, topRight, bottomLeft, bottomCenter, bottomRight
        var id: Self { self }
    }
    var prefix: String
    var suffix: String
    var start: Int64
    var digits: Int
    var position: Position
}

struct FlattenEdit: Encodable, Sendable {
    var fields: [String]?
}

struct AttachmentAdd: Encodable, Sendable {
    var path: String
}

struct SplitPart: Encodable, Sendable, Equatable {
    var pages: [Int]
    var name: String?
}

/// An output beside the document, which is unchanged; exactly one field is set.
struct WriteRequest: Encodable, Sendable {
    struct Split: Encodable, Sendable {
        var directory: String
        var pagesPerPart: Int?
        var parts: [SplitPart]?
    }
    struct AttachmentSave: Encodable, Sendable {
        var fileName: String
        var path: String
    }
    struct Export: Encodable, Sendable {
        var format: String
        var path: String
    }
    var split: Split?
    var attachment: AttachmentSave?
    var export: Export?
}

struct WriteResult: Decodable, Sendable {
    let revision: Int64
    let files: [String]
}

extension EngineDocument {
    func structure() async throws -> DocumentStructure {
        let json: String = try await withEngineCancellation { cancellation in
            var json: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            try check(officeimo_document_structure_json(handle, cancellation, &json, &error), error)
            return takeString(json) ?? "{}"
        }
        return try JSONDecoder().decode(DocumentStructure.self, from: Data(json.utf8))
    }

    /// Applies one edit as one undoable change.
    func edit(_ request: EditRequest) async throws {
        let json = String(decoding: try JSONEncoder().encode(request), as: UTF8.self)
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            try check(json.withCString { officeimo_document_edit_json(handle, $0, cancellation, &error) }, error)
        }
    }

    /// Writes an output beside the document and returns the files written, in order.
    func write(_ request: WriteRequest) async throws -> [URL] {
        let body = String(decoding: try JSONEncoder().encode(request), as: UTF8.self)
        let json: String = try await withEngineCancellation { cancellation in
            var json: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            try check(body.withCString { officeimo_document_write_json(handle, $0, cancellation, &json, &error) }, error)
            return takeString(json) ?? "{}"
        }
        return try JSONDecoder().decode(WriteResult.self, from: Data(json.utf8)).files.map { URL(filePath: $0) }
    }
}
