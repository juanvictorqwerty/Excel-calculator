import 'dart:io';
import 'package:excel/excel.dart';
import 'package:path/path.dart' as path;
import 'gpa_calculator.dart';

/// Exception for Excel parsing errors
class ExcelParseException implements Exception {
  final String message;
  ExcelParseException(this.message);
  
  @override
  String toString() => 'ExcelParseException: $message';
}

/// Handles reading and writing Excel files
class ExcelParser {
  /// Reads an Excel file and extracts student data
  /// Returns a list of StudentResult objects
  static List<StudentResult> readExcelFile(String filePath) {
    // Check if file exists
    final file = File(filePath);
    if (!file.existsSync()) {
      throw ExcelParseException('File not found: $filePath');
    }

    // Check file extension
    final extension = path.extension(filePath).toLowerCase();
    if (extension != '.xlsx' && extension != '.xls') {
      throw ExcelParseException('Invalid file format. Expected .xlsx or .xls');
    }

    // Read the Excel file
    final bytes = file.readAsBytesSync();
    final excel = Excel.decodeBytes(bytes);

    if (excel.tables.isEmpty) {
      throw ExcelParseException('No sheets found in Excel file');
    }

    // Find the sheet with "Name" column
    final sheet = _findSheetWithNameColumn(excel);
    if (sheet == null) {
      throw ExcelParseException('No sheet with "Name" column found');
    }

    return _parseSheet(sheet);
  }

  /// Finds a sheet that contains a "Name" column
  static Sheet? _findSheetWithNameColumn(Excel excel) {
    for (final tableName in excel.tables.keys) {
      final sheet = excel.tables[tableName];
      if (sheet == null) continue;

      // Check first row for "Name" column
      if (sheet.rows.isNotEmpty) {
        final firstRow = sheet.rows.first;
        for (final cell in firstRow) {
          final value = _extractCellValue(cell?.value);
          if (value != null && 
              value.toString().toLowerCase().trim() == 'name') {
            return sheet;
          }
        }
      }
    }
    return null;
  }

  /// Parses a sheet and extracts student results
  static List<StudentResult> _parseSheet(Sheet sheet) {
    final results = <StudentResult>[];
    
    if (sheet.rows.length < 2) {
      throw ExcelParseException('Sheet must have at least a header row and one data row');
    }

    final rows = sheet.rows;
    final headerRow = rows.first;
    
    // Find the Name column index and subject column indices
    int? nameColumnIndex;
    final subjectColumnIndices = <int>[];
    
    for (int i = 0; i < headerRow.length; i++) {
      final cell = headerRow[i];
      final value = _extractCellValue(cell?.value)?.toString().trim() ?? '';
      
      if (value.toLowerCase() == 'name') {
        nameColumnIndex = i;
      } else if (value.isNotEmpty) {
        // All other non-empty columns are treated as subjects
        subjectColumnIndices.add(i);
      }
    }

    if (nameColumnIndex == null) {
      throw ExcelParseException('Name column not found');
    }

    if (subjectColumnIndices.isEmpty) {
      throw ExcelParseException('No subject columns found');
    }

    // Parse data rows
    for (int i = 1; i < rows.length; i++) {
      final row = rows[i];
      if (row.isEmpty) continue;

      // Get student name
      final nameCell = row[nameColumnIndex];
      final name = _extractCellValue(nameCell?.value)?.toString().trim() ?? '';
      
      if (name.isEmpty) continue; // Skip empty rows

      // Collect scores from subject columns
      final scores = <double>[];
      for (final colIndex in subjectColumnIndices) {
        if (colIndex < row.length) {
          final cell = row[colIndex];
          final score = _parseScore(_extractCellValue(cell?.value));
          if (score != null) {
            scores.add(score);
          }
        }
      }

      if (scores.isNotEmpty) {
        final average = GpaCalculator.calculateAverage(scores);
        final gpa = GpaCalculator.calculateGpa(scores);
        final grade = GpaCalculator.scoreToGrade(average);
        
        results.add(StudentResult(
          name: name,
          average: average,
          gpa: gpa,
          grade: grade,
        ));
      }
    }

    return results;
  }

  /// Extracts the raw value from a CellValue object
  static dynamic _extractCellValue(CellValue? cellValue) {
    if (cellValue == null) return null;
    
    // Handle different CellValue types
    if (cellValue is TextCellValue) {
      return cellValue.value;
    } else if (cellValue is DoubleCellValue) {
      return cellValue.value;
    } else if (cellValue is IntCellValue) {
      return cellValue.value;
    } else if (cellValue is BoolCellValue) {
      return cellValue.value;
    } else if (cellValue is DateCellValue) {
      return cellValue.toString();
    } else {
      // Fallback to toString if unknown type
      return cellValue.toString();
    }
  }

  /// Parses a cell value into a score (double)
  static double? _parseScore(dynamic value) {
    if (value == null) return null;
    
    if (value is num) {
      return value.toDouble();
    }
    
    if (value is String) {
      final cleaned = value.trim();
      if (cleaned.isEmpty) return null;
      return double.tryParse(cleaned);
    }
    
    return null;
  }

  /// Saves student results to a new Excel file
  static String saveResultsToExcel(List<StudentResult> results, String inputFilePath) {
    final now = DateTime.now();
    final timestamp = 
        '${now.year}${_twoDigits(now.month)}${_twoDigits(now.day)}_'
        '${_twoDigits(now.hour)}${_twoDigits(now.minute)}${_twoDigits(now.second)}';
    
    final directory = path.dirname(inputFilePath);
    final outputFileName = 'GPA_Results_$timestamp.xlsx';
    final outputPath = path.join(directory, outputFileName);

    // Create a new Excel file
    final excel = Excel.createExcel();
    final sheet = excel['GPA Results'];

    // Add header row
    sheet.appendRow([
      TextCellValue('Name'),
      TextCellValue('Average'),
      TextCellValue('GPA'),
      TextCellValue('Grade'),
    ]);

    // Add data rows
    for (final result in results) {
      sheet.appendRow([
        TextCellValue(result.name),
        TextCellValue(result.average.toStringAsFixed(2)),
        TextCellValue(result.gpa.toStringAsFixed(2)),
        TextCellValue(result.grade),
      ]);
    }

    // Save the file
    final file = File(outputPath);
    final bytes = excel.encode();
    if (bytes != null) {
      file.writeAsBytesSync(bytes);
    } else {
      throw ExcelParseException('Failed to encode Excel file');
    }

    return outputPath;
  }

  /// Helper to format two-digit numbers
  static String _twoDigits(int n) => n.toString().padLeft(2, '0');

  /// Gets the subject column names from the Excel file
  static List<String> getSubjectNames(String filePath) {
    final file = File(filePath);
    if (!file.existsSync()) {
      throw ExcelParseException('File not found: $filePath');
    }

    final bytes = file.readAsBytesSync();
    final excel = Excel.decodeBytes(bytes);

    final sheet = _findSheetWithNameColumn(excel);
    if (sheet == null) {
      throw ExcelParseException('No sheet with "Name" column found');
    }

    if (sheet.rows.isEmpty) {
      return [];
    }

    final headerRow = sheet.rows.first;
    final subjects = <String>[];
    
    for (final cell in headerRow) {
      final value = _extractCellValue(cell?.value)?.toString().trim() ?? '';
      if (value.isNotEmpty && value.toLowerCase() != 'name') {
        subjects.add(value);
      }
    }

    return subjects;
  }
}