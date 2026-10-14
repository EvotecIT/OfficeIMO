import SwiftUI

/// The document's certificate signatures: who signed, when, and whether the document is unchanged since.
struct SignaturesSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var report: CheckResult.Signatures?
    @State private var problem: String?

    var body: some View {
        ReportForm(title: "Signatures", problem: problem, isLoading: report == nil && problem == nil, done: { dismiss() }) {
            if let report {
                if !report.hasSignatures {
                    Section {
                        Label("This document has no certificate signatures.", systemImage: "signature")
                            .foregroundStyle(.secondary)
                    }
                }
                ForEach(report.signatures) { signature in
                    Section {
                        Label(Self.statusText(signature), systemImage: Self.statusSymbol(signature.status))
                            .foregroundStyle(Self.statusColor(signature.status))
                        if let time = signature.signingTime.flatMap({ try? Date($0, strategy: .iso8601) }) {
                            LabeledContent("Signed", value: time, format: .dateTime)
                        } else if let time = signature.signingTime {
                            LabeledContent("Signed", value: time)
                        }
                        if let reason = signature.reason, !reason.isEmpty { LabeledContent("Reason", value: reason) }
                        if let location = signature.location, !location.isEmpty { LabeledContent("Location", value: location) }
                        LabeledContent("Covers", value: signature.coversWholeDocument
                                       ? String(localized: "The whole document") : String(localized: "An earlier version"))
                        ForEach(signature.findings, id: \.self) { Text($0).font(.callout).foregroundStyle(.secondary) }
                    } header: {
                        Text(signature.signer ?? signature.field ?? String(localized: "Signature"))
                    }
                }
                if !report.findings.isEmpty {
                    Section("Findings") {
                        ForEach(report.findings, id: \.self) { Text($0).font(.callout) }
                    }
                }
            }
        }
        .task {
            do { report = try await session.check(CheckRequest(signatures: true)).signatures }
            catch { problem = error.localizedDescription }
        }
    }

    static func statusText(_ signature: CheckResult.Signature) -> LocalizedStringKey {
        switch signature.status {
        case "valid": "Valid, and the document hasn’t changed since it was signed"
        case "intact": "The document hasn’t changed since it was signed; the certificate isn’t trusted on this device"
        case "invalid": "Not valid: the document changed after signing, or the signature is damaged"
        case "unchecked": "The signature couldn’t be checked"
        default: "The signature couldn’t be fully checked"
        }
    }

    static func statusSymbol(_ status: String) -> String {
        switch status {
        case "valid": "checkmark.seal.fill"
        case "intact": "checkmark.seal"
        case "invalid": "xmark.seal.fill"
        default: "questionmark.diamond"
        }
    }

    static func statusColor(_ status: String) -> Color {
        switch status {
        case "valid": .green
        case "intact": .primary
        case "invalid": .red
        default: .orange
        }
    }
}

/// Whether the document meets an archival, accessibility or print standard, requirement by requirement.
struct StandardSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var standard = PDFStandard.pdfa2b
    @State private var report: CheckResult.Standard?
    @State private var problem: String?
    @State private var isChecking = false

    var body: some View {
        ReportForm(title: "Check Standard", problem: problem, isLoading: isChecking, done: { dismiss() }) {
            Section {
                Picker("Standard", selection: $standard) {
                    ForEach(PDFStandard.allCases) { Text($0.title).tag($0) }
                }
            } footer: {
                Text("PDF/A is for archiving, PDF/UA for accessibility and PDF/X for printing.")
            }
            if let report {
                Section {
                    Label(report.isReady ? "Meets \(report.name)" : "Doesn’t meet \(report.name) yet",
                          systemImage: report.isReady ? "checkmark.circle.fill" : "exclamationmark.triangle.fill")
                        .foregroundStyle(report.isReady ? .green : .orange)
                }
                Section("Requirements") {
                    ForEach(report.requirements) { requirement in
                        VStack(alignment: .leading, spacing: 3) {
                            Label(requirement.name, systemImage: requirement.status == "met" ? "checkmark.circle" : requirement.status == "missing" ? "xmark.circle" : "questionmark.circle")
                                .foregroundStyle(requirement.status == "met" ? Color.primary : requirement.status == "missing" ? .red : .orange)
                            if requirement.status != "met", !requirement.detail.isEmpty {
                                Text(requirement.detail).font(.callout).foregroundStyle(.secondary)
                            }
                        }
                    }
                }
            }
        }
        .task(id: standard) {
            isChecking = true
            problem = nil
            defer { isChecking = false }
            do { report = try await session.check(CheckRequest(standard: standard.rawValue)).standard }
            catch { problem = error.localizedDescription }
        }
    }
}

/// Document Health: what the document's structure says about it, and copies that are optimized, repaired or sanitized.
struct HealthSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var health: CheckResult.Health?
    @State private var problem: String?
    @State private var profile = "balanced"
    @State private var writing: HealthCopyKind?

    var body: some View {
        ReportForm(title: "Document Health", problem: problem, isLoading: health == nil && problem == nil, done: { dismiss() }) {
            if let health {
                Section("Document") {
                    LabeledContent("File Size", value: Measurement(value: Double(health.fileSize), unit: UnitInformationStorage.bytes),
                                   format: .byteCount(style: .file))
                    LabeledContent("Pages", value: health.pages, format: .number)
                    if let version = health.version { LabeledContent("PDF Version", value: version) }
                    LabeledContent("Accessibility Tags", value: health.isTagged ? String(localized: "Present") : String(localized: "None"))
                    LabeledContent("Scripts and Actions", value: health.hasActiveContent ? String(localized: "Present") : String(localized: "None"))
                    LabeledContent("Attachments", value: health.hasAttachments ? String(localized: "Present") : String(localized: "None"))
                    LabeledContent("Password", value: health.isEncrypted ? String(localized: "Protected") : String(localized: "None"))
                    LabeledContent("Signatures", value: health.isSigned ? String(localized: "Signed") : String(localized: "None"))
                }
                Section {
                    if health.recoveredProblems == 0 && health.unrecoverableProblems == 0 {
                        Label("No structural problems found.", systemImage: "checkmark.circle").foregroundStyle(.green)
                    } else {
                        Label("\(health.recoveredProblems) problem(s) were worked around while reading; \(health.unrecoverableProblems) couldn’t be.",
                              systemImage: "wrench.and.screwdriver")
                        ForEach(health.diagnostics.prefix(8), id: \.self) { Text($0).font(.callout).foregroundStyle(.secondary) }
                    }
                } header: {
                    Text("Structure")
                }
                Section {
                    Picker("Optimize For", selection: $profile) {
                        Text("Balanced").tag("balanced")
                        Text("Smallest File").tag("smallest")
                        Text("Web").tag("web")
                        Text("Archiving").tag("archival")
                    }
                    copyButton(.optimize, "Save Optimized Copy…", "arrow.down.right.and.arrow.up.left")
                    copyButton(.repair, "Save Repaired Copy…", "wrench.and.screwdriver")
                        .disabled(health.recoveredProblems == 0)
                    copyButton(.sanitize, "Save Sanitized Copy…", "shield.lefthalf.filled")
                } header: {
                    Text("Copies")
                } footer: {
                    Text("Optimizing makes the file smaller without changing what it shows. Repairing writes out the problems worked around while reading. Sanitizing removes scripts, actions, attachments and other hidden content. This document stays as it is.")
                }
            }
        }
        .disabled(writing != nil)
        .task {
            do { health = try await session.check(CheckRequest(health: true)).health }
            catch { problem = error.localizedDescription }
        }
    }

    private func copyButton(_ kind: HealthCopyKind, _ title: LocalizedStringKey, _ symbol: String) -> some View {
        Button {
            writing = kind
            problem = nil
            Task {
                defer { writing = nil }
                do {
                    _ = try await session.writeHealthCopy(kind, profile: kind == .optimize ? profile : nil)
                    dismiss()
                } catch {
                    problem = error.localizedDescription
                }
            }
        } label: {
            HStack {
                Label(title, systemImage: symbol)
                if writing == kind { Spacer(); ProgressView().controlSize(.small) }
            }
        }
    }
}

/// Redact: find text, choose the matches, and remove them for good, checked afterwards.
struct RedactSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var find = RedactionFind(text: "")
    @State private var searched: RedactionSearchBinding?
    @State private var marks: [RedactionMark] = []
    @State private var included: Set<Int> = []
    @State private var sanitize = false
    @State private var problem: String?
    @State private var isSearching = false

    var body: some View {
        SheetForm(title: "Redact", confirm: "Redact", isComplete: searched?.matches(find, revision: session.revision) == true && !included.isEmpty, isWriting: isSearching,
                  problem: problem, cancel: { dismiss() }, confirmAction: apply) {
            Section {
                TextField("Text to Remove", text: $find.text)
                    .onSubmit(search)
                Toggle("Match Case", isOn: $find.matchCase)
                Toggle("Regular Expression", isOn: $find.regex)
                Button("Find", systemImage: "magnifyingglass", action: search)
                    .disabled(find.text.trimmingCharacters(in: .whitespaces).isEmpty)
            } footer: {
                Text("Redaction removes the text, images and drawings under each match and covers it in black. Select text on a page to redact just that, or use Mark Areas to Redact to drag across images and drawings.")
            }
            if searched != nil {
                Section {
                    if marks.isEmpty {
                        Text("No matches.").foregroundStyle(.secondary)
                    }
                    ForEach(Array(marks.enumerated()), id: \.offset) { index, mark in
                        Toggle(isOn: Binding { included.contains(index) } set: { if $0 { included.insert(index) } else { included.remove(index) } }) {
                            VStack(alignment: .leading, spacing: 2) {
                                Text(mark.preview ?? find.text).lineLimit(2)
                                Text("Page \(mark.page)").font(.caption).foregroundStyle(.secondary)
                            }
                        }
                    }
                } header: {
                    Text("Matches (\(marks.count))")
                }
                Section {
                    Toggle("Also Remove Hidden Information", isOn: $sanitize)
                } footer: {
                    Text("Removes the document’s properties, attachments, scripts and comments too.")
                }
            }
        }
        .onChange(of: find) {
            if searched?.matches(find, revision: session.revision) != true { searched = nil }
        }
        .onChange(of: session.revision) { searched = nil }
    }

    private func search() {
        let request = RedactionSearchBinding(find: find, revision: session.revision)
        guard !isSearching, !request.find.text.isEmpty else { return }
        problem = nil
        searched = nil
        isSearching = true
        Task {
            defer { isSearching = false }
            do {
                let found = try await session.findRedactions(request.find)
                guard request.matches(find, revision: session.revision) else { return }
                marks = found
                included = Set(marks.indices)
                searched = request
            } catch {
                problem = error.localizedDescription
            }
        }
    }

    private func apply() {
        guard let searched, searched.matches(find, revision: session.revision) else { return }
        session.redact(.init(find: searched.find, include: included.sorted(), removedText: searched.find.regex ? nil : searched.find.text, sanitize: sanitize))
        dismiss()
    }
}

/// Add Link: the selected text goes to a web or mail address, or to a page of this document.
struct LinkSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var toPage = false
    @State private var address = "https://"
    @State private var page = 1

    private var isComplete: Bool {
        toPage ? (1...max(session.pageCount, 1)).contains(page)
            : URL(string: address.trimmingCharacters(in: .whitespaces)).map { ["https", "http", "mailto"].contains($0.scheme?.lowercased() ?? "") && address.count > 8 } ?? false
    }

    var body: some View {
        SheetForm(title: "Add Link", confirm: "Add", isComplete: isComplete, isWriting: false, problem: nil,
                  cancel: { dismiss() }, confirmAction: add) {
            Section {
                Picker("Link To", selection: $toPage) {
                    Text("Web or Email Address").tag(false)
                    Text("Page in This Document").tag(true)
                }
                .pickerStyle(.inline)
                .labelsHidden()
                if toPage {
                    Stepper("Page \(page)", value: $page, in: 1...max(session.pageCount, 1))
                } else {
                    TextField("Address", text: $address)
                        .textContentType(.URL)
                        #if os(iOS)
                        .keyboardType(.URL)
                        .textInputAutocapitalization(.never)
                        #endif
                        .autocorrectionDisabled()
                }
            } footer: {
                if let text = session.selectedText {
                    Text("“\(String(text.prefix(80)))” becomes a link.")
                }
            }
        }
    }

    private func add() {
        session.linkSelection(uri: toPage ? nil : address.trimmingCharacters(in: .whitespaces), page: toPage ? page : nil)
        dismiss()
    }
}

/// A report in a sheet: a grouped form with Done, a progress indicator while it loads, and what went wrong.
private struct ReportForm<Content: View>: View {
    let title: LocalizedStringKey
    let problem: String?
    let isLoading: Bool
    let done: () -> Void
    @ViewBuilder let content: Content

    var body: some View {
        NavigationStack {
            Form {
                content
                if isLoading {
                    Section { ProgressView().frame(maxWidth: .infinity) }
                }
                if let problem {
                    Section {
                        Label(problem, systemImage: "exclamationmark.triangle.fill").foregroundStyle(.red)
                    }
                }
            }
            .formStyle(.grouped)
            .navigationTitle(title)
            #if os(iOS)
            .navigationBarTitleDisplayMode(.inline)
            #endif
            .toolbar {
                ToolbarItem(placement: .confirmationAction) {
                    Button("Done", action: done)
                }
            }
        }
        #if os(macOS)
        .frame(minWidth: 460, minHeight: 420)
        #endif
    }
}
