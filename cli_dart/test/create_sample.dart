import 'dart:io';
import 'package:excel/excel.dart';

void main() {
  // Create a sample Excel file for testing
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];

  // Add header row
  sheet.appendRow([
    TextCellValue('Name'),
    TextCellValue('Math'),
    TextCellValue('Physics'),
    TextCellValue('Chemistry'),
    TextCellValue('Biology'),
    TextCellValue('English'),
  ]);

  // Add sample student data
  final students = [
    ['Alice Johnson', 92.5, 88.0, 90.5, 85.0, 91.0],
    ['Bob Smith', 78.5, 82.0, 75.5, 80.0, 77.0],
    ['Charlie Brown', 65.0, 70.5, 68.0, 72.5, 66.0],
    ['Diana Prince', 95.0, 94.5, 96.0, 93.5, 97.0],
    ['Eve Davis', 55.0, 60.5, 58.0, 62.5, 59.0],
    ['Frank Wilson', 82.0, 85.5, 80.0, 78.5, 84.0],
    ['Grace Lee', 71.5, 68.0, 74.5, 70.0, 73.0],
    ['Henry Taylor', 45.0, 50.5, 48.0, 52.5, 47.0],
  ];

  for (final student in students) {
    sheet.appendRow([
      TextCellValue(student[0] as String),
      DoubleCellValue(student[1] as double),
      DoubleCellValue(student[2] as double),
      DoubleCellValue(student[3] as double),
      DoubleCellValue(student[4] as double),
      DoubleCellValue(student[5] as double),
    ]);
  }

  // Save the file
  final file = File('sample_grades.xlsx');
  final bytes = excel.encode();
  if (bytes != null) {
    file.writeAsBytesSync(bytes);
    print('Sample file created: ${file.absolute.path}');
  } else {
    print('Failed to create sample file');
  }
}