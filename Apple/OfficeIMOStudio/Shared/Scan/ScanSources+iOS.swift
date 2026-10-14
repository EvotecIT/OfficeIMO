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
        guard let images = try await CameraSession().run(), !images.isEmpty else { return nil }
        let output = try await ScanBuilder.makePDF(from: images, named: ScanBuilder.newName)
        StudioTemporaryFiles.markReturnedOutput(output)
        return output
    }

    /// A new searchable PDF from photos, one page each in the order chosen.
    static func newDocumentFromPhotos() async throws -> URL? {
        guard let images = try await PhotosSession().run(), !images.isEmpty else { return nil }
        let output = try await ScanBuilder.makePDF(from: images, named: ScanBuilder.newName)
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
    private var continuation: CheckedContinuation<[CGImage]?, any Error>?
    private var keepAlive: CameraSession?

    func run() async throws -> [CGImage]? {
        try await withCheckedThrowingContinuation { continuation in
            self.continuation = continuation
            let camera = VNDocumentCameraViewController()
            camera.delegate = self
            keepAlive = self
            Task { if !(await ScanSources.present(camera)) { self.finish(.failure(CocoaError(.featureUnsupported)), from: nil) } }
        }
    }

    func documentCameraViewController(_ controller: VNDocumentCameraViewController, didFinishWith scan: VNDocumentCameraScan) {
        do {
            guard scan.pageCount <= ScanBuilder.maximumPages else { throw CocoaError(.fileReadTooLarge) }
            let images = try (0..<scan.pageCount).map { index in
                guard let image = upright(scan.imageOfPage(at: index)) else { throw CocoaError(.fileReadCorruptFile) }
                return image
            }
            finish(.success(images), from: controller)
        } catch { finish(.failure(error), from: controller) }
    }

    func documentCameraViewControllerDidCancel(_ controller: VNDocumentCameraViewController) {
        finish(.success(nil), from: controller)
    }

    func documentCameraViewController(_ controller: VNDocumentCameraViewController, didFailWithError error: any Error) {
        finish(.failure(error), from: controller)
    }

    private func finish(_ result: Result<[CGImage]?, any Error>, from controller: UIViewController?) {
        controller?.dismiss(animated: true)
        continuation?.resume(with: result)
        continuation = nil
        keepAlive = nil
    }
}

@MainActor
private final class PhotosSession: NSObject, PHPickerViewControllerDelegate {
    private var continuation: CheckedContinuation<[CGImage]?, any Error>?
    private var keepAlive: PhotosSession?

    func run() async throws -> [CGImage]? {
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
        let providers = results.map(\.itemProvider)
        Task {
            picker.dismiss(animated: true)
            do {
                guard providers.count <= ScanBuilder.maximumPages else { throw CocoaError(.fileReadTooLarge) }
                var images: [CGImage] = []
                for provider in providers {
                    let image = try await Self.load(provider)
                    guard let upright = upright(image) else { throw CocoaError(.fileReadCorruptFile) }
                    images.append(upright)
                }
                finish(.success(images.isEmpty ? nil : images))
            } catch { finish(.failure(error)) }
        }
    }

    private static func load(_ provider: NSItemProvider) async throws -> UIImage {
        guard provider.canLoadObject(ofClass: UIImage.self) else { throw CocoaError(.fileReadCorruptFile) }
        return try await withCheckedThrowingContinuation { continuation in
            _ = provider.loadObject(ofClass: UIImage.self) { object, error in
                if let error { continuation.resume(throwing: error) }
                else if let image = object as? UIImage { continuation.resume(returning: image) }
                else { continuation.resume(throwing: CocoaError(.fileReadCorruptFile)) }
            }
        }
    }

    private func finish(_ result: Result<[CGImage]?, any Error>) {
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
