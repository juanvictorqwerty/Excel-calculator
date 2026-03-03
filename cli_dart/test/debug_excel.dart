import 'dart:io';
import 'package:excel/excel.dart';

void main() {
  final file = File('sample_grades.xlsx');
  if (!file.existsSync()) {
    print('File not found');
    return;
  }

  final bytes = file.readAsBytesSync();
  final excel = Excel.decodeBytes(bytes);

  print('Sheets: ${excel.tables.keys}');
  
  for (final tableName in excel.tables.keys) {
    final sheet = excel.tables[tableName];
    if (sheet == null) continue;

    print('\nSheet: $tableName');
    print('Rows: ${sheet.rows.length}');
    
    for (int i = 0; i < sheet.rows.length; i++) {
      final row = sheet.rows[i];
      print('Row $i:');
      for (int j = 0; j < row.length; j++) {
        final cell = row[j];
        final value = cell?.value;
        print('  Col $j: $value (type: ${value?.runtimeType})');
      }
    }
  }
}