#if os(iOS)
import PhotosUI
import UIKit
import VisionKit

/// Where a new scan's pages come from on iPad and iPhone: the VisionKit document camera, which finds the page edges,
/// straightens and cleans each page, or photos already taken.
@MainActor
enum ScanSources {
    static var isCameraAvailable: Bool { VNDocumentCameraViewController.isSupported }

    private enum Source {
        case combine, camera, photos, sample, blank
    }

    /// A new PDF from the source chosen on the launch screen's Create PDF.
    static func newDocumentFromChoice() async throws -> URL? {
        switch await chooseSource() {
        case .combine: try await ConversionSources.newDocumentFromFiles()
        case .camera: try await newDocumentFromCamera()
        case .photos: try await newDocumentFromPhotos()
        case .sample: SampleDocument.url
        case .blank: try newBlankDocument()
        case nil: nil
        }
    }

    private static func chooseSource() async -> Source? {
        await withCheckedContinuation { (continuation: CheckedContinuation<Source?, Never>) in
            // An action sheet on an iPad needs an anchor in a view, and the launch screen offers none.
            let style: UIAlertController.Style = UIDevice.current.userInterfaceIdiom == .pad ? .alert : .actionSheet
            let choice = UIAlertController(title: nil, message: nil, preferredStyle: style)
            @MainActor func add(_ title: String.LocalizationValue, _ action: UIAlertAction.Style = .default, _ source: Source?) {
                choice.addAction(UIAlertAction(title: String(localized: title), style: action) { _ in continuation.resume(returning: source) })
            }
            add("Combine Files…", .default, .combine)
            if isCameraAvailable { add("Scan Document", .default, .camera) }
            add("PDF from Photos", .default, .photos)
            add("Try a Sample PDF", .default, .sample)
            add("New Blank PDF", .default, .blank)
            add("Cancel", .cancel, nil)
            Task { if !(await present(choice)) { continuation.resume(returning: nil) } }
        }
    }

    private static func newBlankDocument() throws -> URL {
        let output = StudioTemporaryFiles.root.appending(path: "Scans/\(UUID().uuidString)/\(String(localized: "Untitled")).pdf")
        try FileManager.default.createDirectory(at: output.deletingLastPathComponent(), withIntermediateDirectories: true)
        try StudioDocument.blankPDF().write(to: output)
        StudioTemporaryFiles.markReturnedOutput(output)
        return output
    }

    /// A new searchable PDF from the document camera.
    static func newDocumentFromCamera() async throws -> URL? {
        guard let scan = try await CameraSession().run(), scan.pageCount > 0 else { return nil }
        let output = try await ScanBuilder.makePDF(pageCount: scan.pageCount, named: ScanBuilder.newName) { index in
            guard let image = autoreleasepool(invoking: { upright(scan.imageOfPage(at: index)) }) else {
                throw CocoaError(.fileReadCorruptFile)
            }
            return image
        }
        StudioTemporaryFiles.markReturnedOutput(output)
        return output
    }

    /// A new searchable PDF from photos, one page each in the order chosen.
    static func newDocumentFromPhotos() async throws -> URL? {
        guard let providers = try await PhotosSession().run(), !providers.isEmpty else { return nil }
        let output = try await ScanBuilder.makePDF(pageCount: providers.count, named: ScanBuilder.newName) { index in
            try await ScanBuilder.image(from: providers[index])
        }
        StudioTemporaryFiles.markReturnedOutput(output)
        return output
    }

    /// Presents a picker over the frontmost screen once any sheet or menu on its way out has gone, as UIKit ignores a
    /// presentation that starts during another's dismissal; false when it could not be shown.
    static func present(_ controller: UIViewController) async -> Bool {
        var presenter: UIViewController?
        for _ in 0..<30 {
            // UIKit can attach the presentation after present() returns. Request it once and
            // observe that controller instead of trying to present it again from itself.
            if controller.presentingViewController != nil || presenter?.presentedViewController === controller ||
                controller.viewIfLoaded?.window != nil { return true }
            let scenes = UIApplication.shared.connectedScenes.compactMap { $0 as? UIWindowScene }
            let window = scenes.flatMap(\.windows).first { $0.isKeyWindow } ?? scenes.first?.windows.first
            guard var top = window?.rootViewController else { return false }
            while let presented = top.presentedViewController, !presented.isBeingDismissed { top = presented }
            if presenter == nil, top.presentedViewController == nil, top.transitionCoordinator == nil,
               !top.isBeingDismissed, top.viewIfLoaded?.window != nil {
                presenter = top
                top.present(controller, animated: true)
            }
            try? await Task.sleep(for: .milliseconds(100))
        }
        return false
    }
}

@MainActor
private final class CameraSession: NSObject, @preconcurrency VNDocumentCameraViewControllerDelegate {
    private var continuation: CheckedContinuation<Void, any Error>?
    private var keepAlive: CameraSession?
    private var scan: VNDocumentCameraScan?

    func run() async throws -> VNDocumentCameraScan? {
        try await withCheckedThrowingContinuation { (continuation: CheckedContinuation<Void, any Error>) in
            self.continuation = continuation
            let camera = VNDocumentCameraViewController()
            camera.delegate = self
            keepAlive = self
            Task { if !(await ScanSources.present(camera)) { self.finish(.failure(CocoaError(.featureUnsupported)), from: nil) } }
        }
        return scan
    }

    func documentCameraViewController(_ controller: VNDocumentCameraViewController, didFinishWith scan: VNDocumentCameraScan) {
        finish(.success(scan), from: controller)
    }

    func documentCameraViewControllerDidCancel(_ controller: VNDocumentCameraViewController) {
        finish(.success(nil), from: controller)
    }

    func documentCameraViewController(_ controller: VNDocumentCameraViewController, didFailWithError error: any Error) {
        finish(.failure(error), from: controller)
    }

    private func finish(_ result: Result<VNDocumentCameraScan?, any Error>, from controller: UIViewController?) {
        controller?.dismiss(animated: true)
        // UIKit values stay on this actor; only completion crosses the continuation.
        switch result {
        case .success(let scan): self.scan = scan; continuation?.resume()
        case .failure(let error): continuation?.resume(throwing: error)
        }
        continuation = nil
        keepAlive = nil
    }
}

@MainActor
private final class PhotosSession: NSObject, PHPickerViewControllerDelegate {
    private var continuation: CheckedContinuation<Void, any Error>?
    private var keepAlive: PhotosSession?
    private var providers: [NSItemProvider]?

    func run() async throws -> [NSItemProvider]? {
        try await withCheckedThrowingContinuation { (continuation: CheckedContinuation<Void, any Error>) in
            self.continuation = continuation
            var configuration = PHPickerConfiguration()
            configuration.filter = .images
            configuration.selectionLimit = ScanBuilder.maximumPages
            configuration.selection = .ordered
            let picker = PHPickerViewController(configuration: configuration)
            picker.delegate = self
            keepAlive = self
            Task { if !(await ScanSources.present(picker)) { self.finish(.failure(CocoaError(.featureUnsupported))) } }
        }
        return providers
    }

    func picker(_ picker: PHPickerViewController, didFinishPicking results: [PHPickerResult]) {
        picker.dismiss(animated: true)
        let providers = results.map(\.itemProvider)
        finish(.success(providers.isEmpty ? nil : providers))
    }

    private func finish(_ result: Result<[NSItemProvider]?, any Error>) {
        // Item providers remain main-actor state until ScanBuilder consumes them.
        switch result {
        case .success(let providers): self.providers = providers; continuation?.resume()
        case .failure(let error): continuation?.resume(throwing: error)
        }
        continuation = nil
        keepAlive = nil
    }
}

/// The image drawn upright, as a photo's orientation shows it.
@MainActor
private func upright(_ image: UIImage) -> CGImage? {
    if image.imageOrientation == .up, let cgImage = image.cgImage { return cgImage }
    let format = UIGraphicsImageRendererFormat()
    format.scale = 1
    return UIGraphicsImageRenderer(size: CGSize(width: image.size.width * image.scale, height: image.size.height * image.scale), format: format)
        .image { _ in image.draw(in: CGRect(origin: .zero, size: CGSize(width: image.size.width * image.scale, height: image.size.height * image.scale))) }
        .cgImage
}
#endif
