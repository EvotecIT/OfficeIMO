window.officeimoConverter = {
  createObjectUrl(bytes, contentType) {
    const blob = new Blob([bytes], { type: contentType || "application/octet-stream" });
    return URL.createObjectURL(blob);
  },
  revokeObjectUrl(url) {
    if (url) {
      URL.revokeObjectURL(url);
    }
  }
};
