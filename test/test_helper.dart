import 'dart:io';

const testOutputDir = './test/test_output';

void saveTestOutput(List<int>? bytes, String filename) {
  if (bytes == null) return;
  final dir = Directory(testOutputDir);
  if (!dir.existsSync()) dir.createSync(recursive: true);
  File('$testOutputDir/$filename.xlsx').writeAsBytesSync(bytes);
}
