#if os(iOS)
import PhotosUI
import UIKit
import VisionKit

/// Where a new scan's pages come from on iPad and iPhone: the VisionKit document camera, which finds the page edges,
/// straightens and cleans each page, or photos already taken.
@MainActor
enum ScanSources {
    static var isCameraAvailable: Bool { VNDocumentCameraViewController.isSupported }

    /// A new searchable PDF from the document camera, for the launch screen's Scan Document.
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
        for _ in 0..<30 {
            let scenes = UIApplication.shared.connectedScenes.compactMap { $0 as? UIWindowScene }
            let window = scenes.flatMap(\.windows).first { $0.isKeyWindow } ?? scenes.first?.windows.first
            guard var top = window?.rootViewController else { return false }
            while let presented = top.presentedViewController, !presented.isBeingDismissed { top = presented }
            if top.presentedViewController == nil, top.transitionCoordinator == nil {
                top.present(controller, animated: true)
                if controller.presentingViewController != nil { return true }
            }
            try? await Task.sleep(for: .milliseconds(100))
        }
        return false
    }
}

@MainActor
private final class CameraSession: NSObject, @preconcurrency VNDocumentCameraViewControllerDelegate {
    private var continuation: CheckedContinuation<VNDocumentCameraScan?, any Error>?
    private var keepAlive: CameraSession?

    func run() async throws -> VNDocumentCameraScan? {
        try await withCheckedThrowingContinuation { continuation in
            self.continuation = continuation
            let camera = VNDocumentCameraViewController()
            camera.delegate = self
            keepAlive = self
            Task { if !(await ScanSources.present(camera)) { self.finish(.failure(CocoaError(.featureUnsupported)), from: nil) } }
        }
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
        continuation?.resume(with: result)
        continuation = nil
        keepAlive = nil
    }
}

@MainActor
private final class PhotosSession: NSObject, PHPickerViewControllerDelegate {
    private var continuation: CheckedContinuation<[NSItemProvider]?, any Error>?
    private var keepAlive: PhotosSession?

    func run() async throws -> [NSItemProvider]? {
        try await withCheckedThrowingContinuation { continuation in
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
    }

    func picker(_ picker: PHPickerViewController, didFinishPicking results: [PHPickerResult]) {
        picker.dismiss(animated: true)
        let providers = results.map(\.itemProvider)
        finish(.success(providers.isEmpty ? nil : providers))
    }

    private func finish(_ result: Result<[NSItemProvider]?, any Error>) {
        continuation?.resume(with: result)
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
