import SwiftUI

/// The document's facts from the engine, in the system's grouped form.
struct DocumentInspector: View {
    let session: DocumentSession

    var body: some View {
        Form {
            if let facts = session.facts {
                Section("Document") {
                    LabeledContent("Name", value: facts.fileName)
                    LabeledContent("Pages", value: facts.pages.count, format: .number)
                    LabeledContent("File Size", value: Measurement(value: Double(facts.fileSize), unit: UnitInformationStorage.bytes),
                                   format: .byteCount(style: .file))
                }
                if facts.pages.indices.contains(session.currentPage - 1) {
                    let page = facts.pages[session.currentPage - 1]
                    Section("Page \(page.number)") {
                        LabeledContent("Size", value: Self.paperSize(page.displaySize))
                        LabeledContent("Rotation", value: Measurement(value: Double(page.rotation), unit: UnitAngle.degrees),
                                       format: .measurement(width: .narrow))
                    }
                }
                Section("Permissions") {
                    LabeledContent("Copy Text", value: facts.canExtractText ? "Allowed" : "Not Allowed")
                    LabeledContent("Extract Content", value: facts.canExtractContent ? "Allowed" : "Not Allowed")
                }
            }
        }
        .formStyle(.grouped)
    }

    /// The size in the reader's units: millimetres, or inches where the US measurement system is used.
    static func paperSize(_ size: CGSize) -> String {
        let metric = Locale.current.measurementSystem != .us
        let unit = metric ? UnitLength.millimeters : UnitLength.inches
        let number = FloatingPointFormatStyle<Double>.number.precision(.fractionLength(0...(metric ? 0 : 2)))
        let width = Measurement(value: size.width / 72, unit: UnitLength.inches).converted(to: unit)
        let height = Measurement(value: size.height / 72, unit: UnitLength.inches).converted(to: unit)
        return "\(width.value.formatted(number)) × \(height.formatted(.measurement(width: .abbreviated, usage: .asProvided, numberFormatStyle: number)))"
    }
}
