import SwiftUI
import PDFKit

/// Native PDFKit supplies selection, scrolling, zoom and accessibility; OfficeIMO performs edits.
struct PDFViewport {
    let document: PDFDocument
    @Binding var page: Int

    func makeCoordinator() -> Coordinator { Coordinator(page: $page) }
    func configure(_ view: PDFView, coordinator: Coordinator) {
        #if os(macOS)
        view.backgroundColor = .underPageBackgroundColor
        #else
        view.backgroundColor = .secondarySystemBackground
        #endif
        view.autoScales = true
        view.displayMode = .singlePageContinuous
        view.displayDirection = .vertical
        view.displaysPageBreaks = true
        coordinator.observe(view)
    }
    func update(_ view: PDFView, coordinator: Coordinator) {
        coordinator.page = $page
        let changed = view.document !== document
        if changed { view.document = document }
        let displayed = view.currentPage.map { document.index(for: $0) }
        if changed || displayed != page, let target = document.page(at: page) { view.go(to: target) }
    }

    final class Coordinator: NSObject {
        var page: Binding<Int>
        private var observation: NSObjectProtocol?
        init(page: Binding<Int>) { self.page = page }
        func observe(_ view: PDFView) {
            observation = NotificationCenter.default.addObserver(forName: .PDFViewPageChanged, object: view, queue: .main) { [weak self, weak view] _ in
                guard let self, let view, let current = view.currentPage, let document = view.document else { return }
                let index = document.index(for: current)
                guard index != NSNotFound else { return }
                DispatchQueue.main.async { [weak self, weak view, weak document] in
                    guard let view, let document, view.document === document,
                          view.currentPage === current else { return }
                    self?.page.wrappedValue = index
                }
            }
        }
        deinit { if let observation { NotificationCenter.default.removeObserver(observation) } }
    }
}

#if os(macOS)
extension PDFViewport: NSViewRepresentable {
    func makeNSView(context: Context) -> PDFView {
        let view = PDFView(); configure(view, coordinator: context.coordinator); return view
    }
    func updateNSView(_ view: PDFView, context: Context) { update(view, coordinator: context.coordinator) }
}
#else
extension PDFViewport: UIViewRepresentable {
    func makeUIView(context: Context) -> PDFView {
        let view = PDFView(); configure(view, coordinator: context.coordinator); return view
    }
    func updateUIView(_ view: PDFView, context: Context) { update(view, coordinator: context.coordinator) }
}
#endif
