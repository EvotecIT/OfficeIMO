import Foundation
import XCTest
@testable import OfficeIMOStudioState

/// Files and Finder offer the app for a type only when its Info.plist declares it, so both platforms' declarations
/// must match the formats the viewer opens.
final class ViewerFormatsTests: XCTestCase {
    private static let appFolder = URL(fileURLWithPath: #filePath).deletingLastPathComponent().deletingLastPathComponent()
        .appending(path: "OfficeIMOStudio")
    private static let infoPlists = ["Info.plist", "macOS/Info.plist"].map { appFolder.appending(path: $0) }

    private func plist(_ url: URL) throws -> [String: Any] {
        let data = try Data(contentsOf: url)
        return try XCTUnwrap(PropertyListSerialization.propertyList(from: data, format: nil) as? [String: Any])
    }

    func testBothPlatformsOfferStudioAsViewerForEveryViewedFormat() throws {
        for url in Self.infoPlists {
            let types = try XCTUnwrap(try plist(url)["CFBundleDocumentTypes"] as? [[String: Any]], url.path)
            let viewers = types.filter { $0["CFBundleTypeRole"] as? String == "Viewer" }
            XCTAssertEqual(viewers.count, 1, url.path)
            XCTAssertEqual(viewers.first?["LSHandlerRank"] as? String, "Alternate", url.path)
            XCTAssertEqual(viewers.first?["LSItemContentTypes"] as? [String], ViewerFormats.identifiers, url.path)

            // PDFs stay the only documents the editor opens, so a Word file never opens as a PDF window.
            let editors = types.filter { $0["CFBundleTypeRole"] as? String == "Editor" }
            XCTAssertEqual(editors.flatMap { $0["LSItemContentTypes"] as? [String] ?? [] }, ["com.adobe.pdf"], url.path)
        }
    }

    func testImportedTypesDeclareTheViewedExtensions() throws {
        let formats = Dictionary(uniqueKeysWithValues: ViewerFormats.formats.map { ($0.identifier, $0.extensions) })
        for url in Self.infoPlists {
            let imports = try XCTUnwrap(try plist(url)["UTImportedTypeDeclarations"] as? [[String: Any]], url.path)
            XCTAssertEqual(Set(imports.compactMap { $0["UTTypeIdentifier"] as? String }), ViewerFormats.importedIdentifiers, url.path)
            for declaration in imports {
                let identifier = try XCTUnwrap(declaration["UTTypeIdentifier"] as? String)
                let tags = declaration["UTTypeTagSpecification"] as? [String: Any]
                XCTAssertEqual(tags?["public.filename-extension"] as? [String], formats[identifier], identifier)
                XCTAssertFalse((declaration["UTTypeConformsTo"] as? [String] ?? []).isEmpty, identifier)
            }
        }
    }

    func testEmptyAndOversizedDocumentsExplainWhyTheyCantBeShown() throws {
        XCTAssertNil(ViewerFormats.sizeProblem(byteCount: 1, name: "Report.docx"))
        XCTAssertNil(ViewerFormats.sizeProblem(byteCount: ViewerFormats.maximumBytes, name: "Report.docx"))
        XCTAssertTrue(try XCTUnwrap(ViewerFormats.sizeProblem(byteCount: 0, name: "Report.docx")).contains("Report.docx"))
        let tooLarge = try XCTUnwrap(ViewerFormats.sizeProblem(byteCount: ViewerFormats.maximumBytes + 1, name: "Deck.pptx"))
        XCTAssertTrue(tooLarge.contains("Deck.pptx"))
        XCTAssertTrue(tooLarge.contains("100 MB"))
    }

    func testEachViewedFormatIsListedOnceWithItsExtensions() {
        XCTAssertEqual(Set(ViewerFormats.identifiers).count, ViewerFormats.formats.count)
        let extensions = ViewerFormats.formats.flatMap(\.extensions)
        XCTAssertEqual(Set(extensions).count, extensions.count)
        XCTAssertTrue(extensions.allSatisfy { !$0.isEmpty && $0 == $0.lowercased() && !$0.hasPrefix(".") })
        XCTAssertTrue(ViewerFormats.importedIdentifiers.isSubset(of: Set(ViewerFormats.identifiers)))
    }
}
