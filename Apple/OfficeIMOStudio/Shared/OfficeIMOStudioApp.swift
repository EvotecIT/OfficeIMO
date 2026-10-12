import SwiftUI

@main
struct OfficeIMOStudioApp: App {
    @State private var model: JourneyModel

    init() {
        let model = JourneyModel()
        _model = State(initialValue: model)
        // Unattended runs start with the app rather than a window, which may never appear when launched from a shell.
        Task { await Autorun.run(model) }
    }

    var body: some Scene {
        WindowGroup {
            JourneyView(model: model)
        }
    }
}

/// Launch arguments for unattended runs:
/// -OfficeIMOJourney runs the journey and writes Results/journey.json;
/// -OfficeIMOMeasure <file> measures Documents/Measurements/<file> and writes Results/measure-<file>.json;
/// -OfficeIMOExit quits afterwards.
@MainActor
enum Autorun {
    static func run(_ model: JourneyModel) async {
        let arguments = ProcessInfo.processInfo.arguments
        await model.start()
        if arguments.contains("-OfficeIMOJourney") {
            await model.runJourney()
        }
        if let index = arguments.firstIndex(of: "-OfficeIMOMeasure"), index + 1 < arguments.count {
            let name = arguments[index + 1]
            do {
                let report = try await MeasurementRunner.run(model.folders.measurements.appending(path: name),
                                                             display: .current, launch: model.launch, engineVersion: model.engineVersion)
                try model.folders.write(report, to: "measure-\(name).json")
            } catch {
                try? model.folders.write(["file": name, "error": error.localizedDescription], to: "measure-\(name).json")
            }
        }
        if arguments.contains("-OfficeIMOExit") {
            exit(0)
        }
    }
}
