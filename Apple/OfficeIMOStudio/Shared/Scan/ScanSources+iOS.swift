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
        guard let images = await CameraSession().run(), !images.isEmpty else { return nil }
        return try await ScanBuilder.makePDF(from: images, named: ScanBuilder.newName)
    }

    /// A new searchable PDF from photos, one page each in the order chosen.
    static func newDocumentFromPhotos() async throws -> URL? {
        guard let images = await PhotosSession().run(), !images.isEmpty else { return nil }
        return try await ScanBuilder.makePDF(from: images, named: ScanBuilder.newName)
    }

    fileprivate static func present(_ controller: UIViewController) -> Bool {
        let scenes = UIApplication.shared.connectedScenes.compactMap { $0 as? UIWindowScene }
        let window = scenes.flatMap(\.windows).first { $0.isKeyWindow } ?? scenes.first?.windows.first
        guard var top = window?.rootViewController else { return false }
        while let presented = top.presentedViewController { top = presented }
        top.present(controller, animated: true)
        return true
    }
}

@MainActor
private final class CameraSession: NSObject, @preconcurrency VNDocumentCameraViewControllerDelegate {
    private var continuation: CheckedContinuation<[CGImage]?, Never>?
    private var keepAlive: CameraSession?

    func run() async -> [CGImage]? {
        await withCheckedContinuation { continuation in
            self.continuation = continuation
            let camera = VNDocumentCameraViewController()
            camera.delegate = self
            keepAlive = self
            if !ScanSources.present(camera) { finish(nil, from: nil) }
        }
    }

    func documentCameraViewController(_ controller: VNDocumentCameraViewController, didFinishWith scan: VNDocumentCameraScan) {
        let images = (0..<scan.pageCount).compactMap { upright(scan.imageOfPage(at: $0)) }
        finish(images, from: controller)
    }

    func documentCameraViewControllerDidCancel(_ controller: VNDocumentCameraViewController) {
        finish(nil, from: controller)
    }

    func documentCameraViewController(_ controller: VNDocumentCameraViewController, didFailWithError error: any Error) {
        finish(nil, from: controller)
    }

    private func finish(_ images: [CGImage]?, from controller: UIViewController?) {
        controller?.dismiss(animated: true)
        continuation?.resume(returning: images)
        continuation = nil
        keepAlive = nil
    }
}

@MainActor
private final class PhotosSession: NSObject, PHPickerViewControllerDelegate {
    private var continuation: CheckedContinuation<[CGImage]?, Never>?
    private var keepAlive: PhotosSession?

    func run() async -> [CGImage]? {
        await withCheckedContinuation { continuation in
            self.continuation = continuation
            var configuration = PHPickerConfiguration()
            configuration.filter = .images
            configuration.selectionLimit = 0
            configuration.selection = .ordered
            let picker = PHPickerViewController(configuration: configuration)
            picker.delegate = self
            keepAlive = self
            if !ScanSources.present(picker) { finish(nil) }
        }
    }

    func picker(_ picker: PHPickerViewController, didFinishPicking results: [PHPickerResult]) {
        let providers = results.map(\.itemProvider)
        Task {
            picker.dismiss(animated: true)
            var images: [CGImage] = []
            for provider in providers where provider.canLoadObject(ofClass: UIImage.self) {
                if let image = await Self.load(provider), let upright = upright(image) { images.append(upright) }
            }
            finish(images.isEmpty ? nil : images)
        }
    }

    private static func load(_ provider: NSItemProvider) async -> UIImage? {
        await withCheckedContinuation { continuation in
            _ = provider.loadObject(ofClass: UIImage.self) { object, _ in
                continuation.resume(returning: object as? UIImage)
            }
        }
    }

    private func finish(_ images: [CGImage]?) {
        continuation?.resume(returning: images)
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
