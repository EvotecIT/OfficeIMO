import Foundation

/// The protection and signing sheets a window shows.
enum SecuritySheet: String, Identifiable {
    case protect, removePassword, sign

    var id: String { rawValue }
}

extension DocumentSession {
    var isEncrypted: Bool { facts?.isEncrypted == true }
    var canChangeSecurity: Bool { isReady && !isWorking }
    var canSignWithCertificate: Bool { isReady && !isWorking && facts?.canSign == true }
    /// Whether changing the protection needs the document's permissions (owner) password.
    var needsOwnerPassword: Bool { facts?.canChangeProtection == false }

    /// The document's name without its extension, for the names of copies.
    var documentBaseName: String {
        (facts?.fileName as NSString?)?.deletingPathExtension ?? String(localized: "Document")
    }

    /// Writes a copy of the document, with a password to open it and the permissions given, for the user to save.
    func exportProtectedCopy(password: String, ownerPassword: String?, currentOwnerPassword: String?,
                             allowPrinting: Bool, allowCopying: Bool, allowChanges: Bool) async throws {
        try await writeCopy(named: String(localized: "\(documentBaseName) – Protected.pdf")) { engine, url in
            try await engine.saveProtectedCopy(ProtectionRequest(path: url.path, userPassword: password, ownerPassword: ownerPassword,
                                                                 currentOwnerPassword: currentOwnerPassword, allowPrinting: allowPrinting,
                                                                 allowCopying: allowCopying, allowChanges: allowChanges))
        }
    }

    /// Writes a copy of the document without its password, for the user to save.
    func exportUnprotectedCopy(currentOwnerPassword: String?) async throws {
        try await writeCopy(named: String(localized: "\(documentBaseName) – Unprotected.pdf")) { engine, url in
            try await engine.saveProtectedCopy(ProtectionRequest(path: url.path, currentOwnerPassword: currentOwnerPassword))
        }
    }

    /// Writes a copy of the document signed with a certificate, for the user to save.
    func exportSignedCopy(certificate: URL, password: String, reason: String?, location: String?) async throws {
        try await writeCopy(named: String(localized: "\(documentBaseName) – Signed.pdf")) { engine, url in
            try await engine.saveSignedCopy(SigningRequest(path: url.path, certificateURL: certificate, certificatePassword: password,
                                                           reason: reason, location: location))
        }
    }

    /// Writes a copy into the app's temporary folder; the window offers it to save once the sheet that asked for it
    /// has closed.
    private func writeCopy(named name: String, write: @escaping @Sendable (EngineDocument, URL) async throws -> Void) async throws {
        guard let engine else { return }
        let folder = FileManager.default.temporaryDirectory.appending(path: "Exported/\(UUID().uuidString)", directoryHint: .isDirectory)
        let url = folder.appending(path: name)
        isWorking = true
        defer { isWorking = false }
        do {
            try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
            try await write(engine, url)
            pendingExport = ExtractedPages(url: url, document: try ExtractedPDF(url: url))
        } catch {
            try? FileManager.default.removeItem(at: folder)
            throw error
        }
    }

    /// Offers the copy a closed sheet wrote for saving.
    func presentPendingExport() {
        guard let pendingExport else { return }
        self.pendingExport = nil
        extractedPages = pendingExport
    }
}
