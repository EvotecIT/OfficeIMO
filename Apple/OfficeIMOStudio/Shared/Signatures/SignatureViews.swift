import SwiftUI

/// The Signature menu: the signatures kept on this device to sign with, a new one and removing one.
struct SignatureMenu: View {
    let session: DocumentSession
    @State private var store = SignatureStore.shared

    var body: some View {
        Menu("Signature", systemImage: "signature") {
            ForEach(store.signatures) { signature in
                Button { session.beginSigning(with: signature) } label: {
                    Label { Text("Signature") } icon: { SignatureImage(signature: signature) }
                }
            }
            Button("New Signature…", systemImage: "plus") { session.isCreatingSignature = true }
            if !store.signatures.isEmpty {
                Menu("Remove Signature", systemImage: "trash") {
                    ForEach(store.signatures) { signature in
                        Button(role: .destructive) { store.remove(signature) } label: {
                            Label { Text("Signature") } icon: { SignatureImage(signature: signature) }
                        }
                    }
                }
            }
        }
        .disabled(!session.canEditAnnotations)
    }
}

/// A signature drawn as an image for menus, which show images rather than views.
private struct SignatureImage: View {
    let signature: Signature

    var body: some View {
        let height = 24.0
        let renderer = ImageRenderer(content: SignatureShape(signature: signature)
            .stroke(.primary, style: StrokeStyle(lineWidth: 1.4, lineCap: .round, lineJoin: .round))
            .frame(width: max(signature.aspect * height, height), height: height)
            .padding(2))
        renderer.scale = 2
        #if os(macOS)
        if let image = renderer.nsImage { return Image(nsImage: image) }
        #else
        if let image = renderer.uiImage { return Image(uiImage: image.withRenderingMode(.alwaysTemplate)) }
        #endif
        return Image(systemName: "signature")
    }
}

/// A signature's strokes fitted to the height of the frame.
struct SignatureShape: Shape {
    let signature: Signature

    func path(in rect: CGRect) -> Path {
        var path = Path()
        let scale = rect.height
        for stroke in signature.strokes {
            guard let first = stroke.first else { continue }
            path.move(to: CGPoint(x: rect.minX + first.x * scale, y: rect.minY + first.y * scale))
            for point in stroke.dropFirst() {
                path.addLine(to: CGPoint(x: rect.minX + point.x * scale, y: rect.minY + point.y * scale))
            }
        }
        return path
    }
}

/// Drawing a new signature with a finger, Apple Pencil, the trackpad or the mouse, as on a sheet of paper.
struct SignaturePad: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var strokes: [[CGPoint]] = []
    @State private var current: [CGPoint] = []

    var body: some View {
        NavigationStack {
            VStack(spacing: 12) {
                ZStack(alignment: .bottom) {
                    RoundedRectangle(cornerRadius: 12)
                        .fill(.white)
                    Rectangle()
                        .fill(.secondary.opacity(0.5))
                        .frame(height: 1)
                        .padding(.horizontal, 24)
                        .padding(.bottom, 56)
                        .accessibilityHidden(true)
                    Canvas { context, _ in
                        var path = Path()
                        for stroke in strokes + [current] where !stroke.isEmpty {
                            path.move(to: stroke[0])
                            for point in stroke.dropFirst() { path.addLine(to: point) }
                        }
                        context.stroke(path, with: .color(Color(red: 0.05, green: 0.1, blue: 0.35)),
                                       style: StrokeStyle(lineWidth: 2.5, lineCap: .round, lineJoin: .round))
                    }
                    .gesture(DragGesture(minimumDistance: 0)
                        .onChanged { current.append($0.location) }
                        .onEnded { _ in
                            strokes.append(current)
                            current = []
                        })
                    .accessibilityLabel("Drawing area")
                }
                .frame(minHeight: 220)
                .environment(\.colorScheme, .light)
                Text("Sign on the line.")
                    .font(.callout)
                    .foregroundStyle(.secondary)
            }
            .padding()
            .navigationTitle("New Signature")
            #if os(iOS)
            .navigationBarTitleDisplayMode(.inline)
            #endif
            .toolbar {
                ToolbarItem(placement: .cancellationAction) {
                    Button("Cancel") { dismiss() }
                }
                ToolbarItem(placement: .destructiveAction) {
                    Button("Clear") { strokes = [] }
                        .disabled(strokes.isEmpty)
                }
                ToolbarItem(placement: .confirmationAction) {
                    Button("Done") {
                        if let signature = Signature(drawn: strokes) {
                            SignatureStore.shared.add(signature)
                            session.beginSigning(with: signature)
                        }
                        dismiss()
                    }
                    .disabled(Signature(drawn: strokes) == nil)
                }
            }
        }
        #if os(macOS)
        .frame(minWidth: 520, minHeight: 360)
        #endif
    }
}

/// Over the page while a signature waits to be placed: where to sign, and Cancel.
struct SigningBanner: View {
    let session: DocumentSession

    var body: some View {
        HStack(spacing: 10) {
            Image(systemName: "signature")
                .accessibilityHidden(true)
            #if os(macOS)
            Text("Click where to sign.")
            #else
            Text("Tap where to sign.")
            #endif
            Button("Cancel") { session.pendingSignature = nil }
                .buttonStyle(.borderless)
        }
        .font(.callout.weight(.medium))
        .padding(.horizontal, 14)
        .padding(.vertical, 8)
        .glassEffect(.regular, in: .capsule)
    }
}
