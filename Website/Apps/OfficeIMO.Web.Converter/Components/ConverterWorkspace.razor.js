export function createObjectUrl(bytes, contentType) {
  const blob = new Blob([bytes], { type: contentType || "application/octet-stream" });
  return URL.createObjectURL(blob);
}

export function revokeObjectUrl(url) {
  if (url) {
    URL.revokeObjectURL(url);
  }
}
