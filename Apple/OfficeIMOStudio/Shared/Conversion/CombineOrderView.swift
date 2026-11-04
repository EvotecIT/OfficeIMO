import SwiftUI

/// The files chosen to combine, in the order their pages will have, before anything is written.
struct CombineOrderView: View {
    @State var files: [URL]
    let finish: ([URL]?) -> Void

    var body: some View {
        NavigationStack {
            List {
                Section {
                    ForEach(files, id: \.self) { file in
                        HStack {
                            Label(file.lastPathComponent, systemImage: "doc")
                                .lineLimit(1)
                                .truncationMode(.middle)
                            #if os(macOS)
                            Spacer()
                            // A list row can be dragged only after a pause, which isn't obvious with a mouse.
                            Button("Move Up", systemImage: "chevron.up") { move(file, by: -1) }
                                .disabled(file == files.first)
                            Button("Move Down", systemImage: "chevron.down") { move(file, by: 1) }
                                .disabled(file == files.last)
                            #endif
                        }
                        #if os(macOS)
                        .buttonStyle(.borderless)
                        .labelStyle(.iconOnly)
                        #endif
                    }
                    .onMove { files.move(fromOffsets: $0, toOffset: $1) }
                } footer: {
                    Text(files.count < 2 ? "Choose at least two files to combine." : "Drag the files into the order you want.")
                }
            }
            #if os(iOS)
            .environment(\.editMode, .constant(.active))
            #endif
            .navigationTitle("Combine Files")
            .toolbar {
                ToolbarItem(placement: .cancellationAction) {
                    Button("Cancel") { finish(nil) }
                }
                ToolbarItem(placement: .confirmationAction) {
                    Button("Combine") { finish(files) }
                        .disabled(files.count < 2)
                }
            }
        }
        #if os(macOS)
        .frame(width: 440, height: 340)
        #endif
    }

    private func move(_ file: URL, by offset: Int) {
        guard let index = files.firstIndex(of: file) else { return }
        files.swapAt(index, index + offset)
    }

    static func sorted(_ files: [URL]) -> [URL] {
        files.sorted { $0.lastPathComponent.localizedStandardCompare($1.lastPathComponent) == .orderedAscending }
    }
}
