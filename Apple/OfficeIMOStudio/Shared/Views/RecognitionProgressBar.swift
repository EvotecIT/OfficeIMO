import SwiftUI

/// Progress of Make Searchable over the pages: the page being read, a bar and Cancel, on glass.
struct RecognitionProgressBar: View {
    let progress: RecognitionProgress
    let cancel: () -> Void

    var body: some View {
        HStack(spacing: 14) {
            VStack(alignment: .leading, spacing: 6) {
                Text(progress.page == 0 ? "Preparing to read text…" : "Reading text on page \(progress.page) of \(progress.pageCount)")
                    .font(.callout)
                    .monospacedDigit()
                    .fixedSize()
                ProgressView(value: Double(max(progress.page - 1, 0)), total: Double(max(progress.pageCount, 1)))
                    .frame(width: 220)
            }
            Button("Cancel", role: .cancel, action: cancel)
        }
        .padding(.horizontal, 18)
        .padding(.vertical, 12)
        .glassEffect(.regular, in: .capsule)
        .padding(.bottom, 24)
        .accessibilityElement(children: .contain)
    }
}
