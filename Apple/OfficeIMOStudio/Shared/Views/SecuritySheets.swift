import SwiftUI
import UniformTypeIdentifiers

/// The sheet for a SecuritySheet: each writes a copy that the window then offers to save.
struct SecuritySheetView: View {
    let session: DocumentSession
    let sheet: SecuritySheet

    var body: some View {
        switch sheet {
        case .protect: ProtectSheet(session: session)
        case .removePassword: RemovePasswordSheet(session: session)
        case .sign: SignSheet(session: session)
        }
    }
}

/// Protect with Password: a password to open the copy and, as Preview's permissions do, what others may do with it.
private struct ProtectSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var password = ""
    @State private var verify = ""
    @State private var allowPrinting = true
    @State private var allowCopying = true
    @State private var allowChanges = true
    @State private var ownerPassword = ""
    @State private var currentOwnerPassword = ""
    @State private var problem: String?
    @State private var isWriting = false

    private var restricts: Bool { !allowPrinting || !allowCopying || !allowChanges }
    private var isComplete: Bool {
        !password.isEmpty && password == verify && (!restricts || (!ownerPassword.isEmpty && ownerPassword != password))
            && (!session.needsOwnerPassword || !currentOwnerPassword.isEmpty)
    }

    var body: some View {
        SecurityForm(title: "Protect with Password", confirm: "Protect", isComplete: isComplete, isWriting: isWriting, problem: problem,
                     cancel: { dismiss() }, confirmAction: protect) {
            Section {
                SecureField("Password", text: $password)
                    .textContentType(.newPassword)
                SecureField("Verify", text: $verify)
                    .textContentType(.newPassword)
            } header: {
                Text("Password to Open")
            } footer: {
                if !verify.isEmpty && verify != password { Text("The passwords don’t match.") }
            }
            Section {
                Toggle("Printing", isOn: $allowPrinting)
                Toggle("Copying Text and Images", isOn: $allowCopying)
                Toggle("Changes, Comments and Form Filling", isOn: $allowChanges)
                if restricts {
                    SecureField("Permissions Password", text: $ownerPassword)
                        .textContentType(.newPassword)
                }
            } header: {
                Text("Allow")
            } footer: {
                Text(restricts
                     ? "Apps that respect permissions ask for the permissions password before allowing what is turned off. It must differ from the password to open."
                     : "Anyone with the password can print, copy and change the copy.")
            }
            if session.needsOwnerPassword {
                CurrentOwnerPasswordSection(password: $currentOwnerPassword)
            }
        }
    }

    private func protect() {
        problem = nil
        isWriting = true
        Task {
            defer { isWriting = false }
            do {
                try await session.exportProtectedCopy(password: password, ownerPassword: restricts ? ownerPassword : nil,
                                                      currentOwnerPassword: session.needsOwnerPassword ? currentOwnerPassword : nil,
                                                      allowPrinting: allowPrinting, allowCopying: allowCopying, allowChanges: allowChanges)
                dismiss()
            } catch {
                problem = SecurityProblem.message(for: error)
            }
        }
    }
}

/// Remove Password: a copy without the password, after the permissions password where the document requires it.
private struct RemovePasswordSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var currentOwnerPassword = ""
    @State private var problem: String?
    @State private var isWriting = false

    var body: some View {
        SecurityForm(title: "Remove Password", confirm: "Remove", isComplete: !session.needsOwnerPassword || !currentOwnerPassword.isEmpty,
                     isWriting: isWriting, problem: problem, cancel: { dismiss() }, confirmAction: remove) {
            Section {
                Text("The copy opens without a password, and anyone can print, copy and change it.")
                    .foregroundStyle(.secondary)
            }
            if session.needsOwnerPassword {
                CurrentOwnerPasswordSection(password: $currentOwnerPassword)
            }
        }
    }

    private func remove() {
        problem = nil
        isWriting = true
        Task {
            defer { isWriting = false }
            do {
                try await session.exportUnprotectedCopy(currentOwnerPassword: session.needsOwnerPassword ? currentOwnerPassword : nil)
                dismiss()
            } catch {
                problem = SecurityProblem.message(for: error)
            }
        }
    }
}

/// Sign with Certificate: a certificate signature from a .p12 or .pfx file, so readers can verify who signed the copy and
/// that it hasn't changed since.
private struct SignSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var certificate: URL?
    @State private var isChoosing = false
    @State private var password = ""
    @State private var reason = ""
    @State private var location = ""
    @State private var problem: String?
    @State private var isWriting = false

    var body: some View {
        SecurityForm(title: "Sign with Certificate", confirm: "Sign", isComplete: certificate != nil, isWriting: isWriting, problem: problem,
                     cancel: { dismiss() }, confirmAction: sign) {
            Section {
                LabeledContent("Certificate") {
                    Button(certificate?.lastPathComponent ?? String(localized: "Choose…")) { isChoosing = true }
                }
                SecureField("Certificate Password", text: $password)
                    .textContentType(.password)
            } footer: {
                Text("A .p12 or .pfx file with your certificate and its private key, such as one exported from Keychain Access or issued by your organization.")
            }
            Section {
                TextField("Reason", text: $reason, prompt: Text("Optional"))
                TextField("Location", text: $location, prompt: Text("Optional"))
                    .textContentType(.addressCity)
            }
        }
        .fileImporter(isPresented: $isChoosing, allowedContentTypes: [.pkcs12]) { result in
            if case .success(let url) = result { certificate = url }
        }
    }

    private func sign() {
        guard let certificate else { return }
        problem = nil
        isWriting = true
        Task {
            defer { isWriting = false }
            do {
                try await session.exportSignedCopy(certificate: certificate, password: password,
                                                   reason: reason.isEmpty ? nil : reason, location: location.isEmpty ? nil : location)
                dismiss()
            } catch {
                problem = SecurityProblem.message(for: error)
            }
        }
    }
}

/// The permissions password of the open document, where its permissions don't allow changing its security.
private struct CurrentOwnerPasswordSection: View {
    @Binding var password: String

    var body: some View {
        Section {
            SecureField("Current Permissions Password", text: $password)
                .textContentType(.password)
        } footer: {
            Text("This document’s permissions don’t allow changing its security without its permissions password.")
        }
    }
}

/// A grouped form with Cancel and a confirming button, as system sheets have, and a problem shown under the fields.
private struct SecurityForm<Content: View>: View {
    let title: LocalizedStringKey
    let confirm: LocalizedStringKey
    let isComplete: Bool
    let isWriting: Bool
    let problem: String?
    let cancel: () -> Void
    let confirmAction: () -> Void
    @ViewBuilder let content: Content

    var body: some View {
        NavigationStack {
            Form {
                content
                if let problem {
                    Section {
                        Label(problem, systemImage: "exclamationmark.triangle.fill")
                            .foregroundStyle(.red)
                    }
                }
            }
            .formStyle(.grouped)
            .navigationTitle(title)
            #if os(iOS)
            .navigationBarTitleDisplayMode(.inline)
            #endif
            .toolbar {
                ToolbarItem(placement: .cancellationAction) {
                    Button("Cancel", action: cancel)
                }
                ToolbarItem(placement: .confirmationAction) {
                    if isWriting {
                        ProgressView().controlSize(.small)
                    } else {
                        Button(confirm, action: confirmAction)
                            .disabled(!isComplete)
                    }
                }
            }
            .disabled(isWriting)
        }
        #if os(macOS)
        .frame(minWidth: 440, minHeight: 360)
        #endif
    }
}

private enum SecurityProblem {
    static func message(for error: any Error) -> String {
        if let error = error as? EngineError, error.status == .invalidPassword {
            return String(localized: "The password is incorrect. Try again.")
        }
        return error.localizedDescription
    }
}

/// Protect with Password, Remove Password and Sign with Certificate, for the File menu and the task menu.
struct SecurityCommands: View {
    let session: DocumentSession

    var body: some View {
        Button("Protect with Password…", systemImage: "lock") { session.securitySheet = .protect }
            .disabled(!session.canChangeSecurity)
        if session.isEncrypted {
            Button("Remove Password…", systemImage: "lock.open") { session.securitySheet = .removePassword }
                .disabled(!session.canChangeSecurity)
        }
        Button("Sign with Certificate…", systemImage: "signature") { session.securitySheet = .sign }
            .disabled(!session.canSignWithCertificate)
    }
}
