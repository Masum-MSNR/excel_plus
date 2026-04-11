part of '../../excel_plus.dart';

Archive _cloneArchive(
  Archive archive,
  Map<String, ArchiveFile> archiveFiles, {
  String? excludedFile,
}) {
  var clone = Archive();
  for (var file in archive.files) {
    if (file.isFile) {
      if (excludedFile != null &&
          file.name.toLowerCase() == excludedFile.toLowerCase()) {
        continue;
      }
      if (archiveFiles.containsKey(file.name)) {
        clone.addFile(archiveFiles[file.name]!);
      } else {
        // Reuse original ArchiveFile reference — no copy needed
        clone.addFile(file);
      }
    }
  }
  return clone;
}
