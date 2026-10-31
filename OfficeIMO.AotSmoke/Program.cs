string root = Path.Combine(Path.GetTempPath(), "officeimo-front-door-native-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try {
    await ConversionBatchScenario.RunAsync(root);
} finally {
    Directory.Delete(root, recursive: true);
}
