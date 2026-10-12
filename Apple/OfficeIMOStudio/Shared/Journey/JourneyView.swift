import SwiftUI

struct JourneyView: View {
    @Bindable var model: JourneyModel
    @Environment(\.displayScale) private var displayScale

    var body: some View {
        NavigationStack {
            VStack(spacing: 0) {
                pageView
                Divider()
                stepList
            }
            .navigationTitle(model.facts?.fileName ?? "OfficeIMO Studio")
            .toolbar { toolbar }
        }
        .onAppear { model.firstViewAppeared() }
    }

    @ViewBuilder private var pageView: some View {
        if let page = model.page {
            Image(decorative: page, scale: displayScale)
                .resizable()
                .aspectRatio(contentMode: .fit)
                .shadow(radius: 2)
                .padding()
                .frame(maxWidth: .infinity, maxHeight: .infinity)
                .accessibilityLabel("Page 1")
        } else {
            ContentUnavailableView {
                Label("No document", systemImage: "doc.richtext")
            } description: {
                Text("Run the journey to open the bundled PDF.")
            } actions: {
                Button("Run Journey") { Task { await model.runJourney() } }
                    .buttonStyle(.borderedProminent)
                    .disabled(model.isBusy)
            }
            .frame(maxWidth: .infinity, maxHeight: .infinity)
        }
    }

    private var stepList: some View {
        List(model.steps) { step in
            HStack(alignment: .firstTextBaseline) {
                Image(systemName: step.passed ? "checkmark.circle.fill" : "xmark.octagon.fill")
                    .foregroundStyle(step.passed ? .green : .red)
                VStack(alignment: .leading) {
                    Text(step.name)
                    Text(step.detail).font(.caption).foregroundStyle(.secondary)
                }
                Spacer()
                Text(step.milliseconds, format: .number.precision(.fractionLength(0))) + Text(" ms")
            }
            .monospacedDigit()
        }
        .frame(minHeight: 180, maxHeight: 260)
    }

    @ToolbarContentBuilder private var toolbar: some ToolbarContent {
        ToolbarItemGroup(placement: .primaryAction) {
            Button("Rotate", systemImage: "rotate.right") { Task { await model.rotate() } }
                .disabled(!model.canEdit)
            Button("Undo", systemImage: "arrow.uturn.backward") { Task { await model.undo() } }
                .disabled(!model.canEdit || model.facts?.canUndo != true)
            Button("Redo", systemImage: "arrow.uturn.forward") { Task { await model.redo() } }
                .disabled(!model.canEdit || model.facts?.canRedo != true)
            Button("Save a Copy", systemImage: "square.and.arrow.down") { Task { await model.saveCopy() } }
                .disabled(!model.canEdit)
            Button("Run Journey", systemImage: "play") { Task { await model.runJourney() } }
                .disabled(model.isBusy)
        }
    }
}
