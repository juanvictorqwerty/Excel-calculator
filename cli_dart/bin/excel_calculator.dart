import 'dart:io';
import 'package:path/path.dart' as path;
import '../lib/excel_parser.dart';
import '../lib/gpa_calculator.dart';

// ANSI color codes for terminal output
const String reset = '\x1B[0m';
const String bold = '\x1B[1m';
const String green = '\x1B[32m';
const String red = '\x1B[31m';
const String cyan = '\x1B[36m';
const String yellow = '\x1B[33m';

void main(List<String> arguments) async {
  printHeader();

  String? filePath;

  // Check for command line argument
  if (arguments.isNotEmpty) {
    filePath = arguments.first;
  } else {
    // Prompt user for file path
    filePath = await promptForFilePath();
  }

  if (filePath == null || filePath.isEmpty) {
    printError('No file path provided.');
    exit(1);
  }

  // Validate and process the file
  try {
    await processExcelFile(filePath);
  } on ExcelParseException catch (e) {
    printError(e.toString());
    exit(1);
  } catch (e) {
    printError('Unexpected error: $e');
    exit(1);
  }
}

/// Prints the application header
void printHeader() {
  print('');
  print('${cyan}╔══════════════════════════════════════════════════════════╗$reset');
  print('${cyan}║$reset           ${bold}Excel GPA Calculator$reset                        ${cyan}║$reset');
  print('${cyan}╚══════════════════════════════════════════════════════════╝$reset');
  print('');
}

/// Prompts the user to enter a file path
Future<String?> promptForFilePath() async {
  stdout.write('${yellow}Enter the path to the Excel file:$reset ');
  final input = stdin.readLineSync();
  return input?.trim();
}

/// Processes the Excel file and displays results
Future<void> processExcelFile(String filePath) async {
  // Normalize the path
  final normalizedPath = path.normalize(filePath);
  final absolutePath = path.isAbsolute(normalizedPath) 
      ? normalizedPath 
      : path.absolute(normalizedPath);

  print('${cyan}Reading file:$reset $absolutePath');
  print('');

  // Read the Excel file
  final results = ExcelParser.readExcelFile(absolutePath);

  if (results.isEmpty) {
    printError('No valid student data found in the file.');
    return;
  }

  // Get subject names for display
  List<String> subjectNames;
  try {
    subjectNames = ExcelParser.getSubjectNames(absolutePath);
  } catch (e) {
    subjectNames = [];
  }

  // Display results
  displayResults(results, subjectNames);

  // Save results to new Excel file
  try {
    final outputPath = ExcelParser.saveResultsToExcel(results, absolutePath);
    print('');
    printSuccess('Results saved to: $outputPath');
  } catch (e) {
    printError('Failed to save results: $e');
  }
}

/// Displays the results in a formatted table
void displayResults(List<StudentResult> results, List<String> subjectNames) {
  print('');
  print('${bold}Found ${results.length} student(s)${reset}');
  if (subjectNames.isNotEmpty) {
    print('Subjects: ${subjectNames.join(", ")}');
  }
  print('');

  // Calculate column widths
  final nameWidth = _calculateNameWidth(results);
  final avgWidth = 10;
  final gpaWidth = 8;
  final gradeWidth = 8;

  // Print table header
  printTableSeparator(nameWidth, avgWidth, gpaWidth, gradeWidth);
  printTableRow('Name', 'Average', 'GPA', 'Grade', 
      nameWidth, avgWidth, gpaWidth, gradeWidth, isHeader: true);
  printTableSeparator(nameWidth, avgWidth, gpaWidth, gradeWidth);

  // Print data rows
  for (final result in results) {
    final coloredGrade = GpaCalculator.getColoredGrade(result.grade);
    final coloredGpa = GpaCalculator.getColoredGpa(result.gpa);
    
    printTableRowWithColors(
      result.name,
      result.average.toStringAsFixed(2),
      coloredGpa,
      coloredGrade,
      nameWidth,
      avgWidth,
      gpaWidth,
      gradeWidth,
    );
  }

  printTableSeparator(nameWidth, avgWidth, gpaWidth, gradeWidth);

  // Print summary statistics
  print('');
  printSummary(results);
}

/// Calculates the width needed for the name column
int _calculateNameWidth(List<StudentResult> results) {
  int maxLen = 4; // Minimum width for "Name"
  for (final result in results) {
    if (result.name.length > maxLen) {
      maxLen = result.name.length;
    }
  }
  return maxLen + 2; // Add padding
}

/// Prints a table separator line
void printTableSeparator(int nameWidth, int avgWidth, int gpaWidth, int gradeWidth) {
  final buffer = StringBuffer();
  buffer.write('+-');
  buffer.write('-' * nameWidth);
  buffer.write('-+-');
  buffer.write('-' * avgWidth);
  buffer.write('-+-');
  buffer.write('-' * gpaWidth);
  buffer.write('-+-');
  buffer.write('-' * gradeWidth);
  buffer.write('-+');
  print(buffer.toString());
}

/// Prints a table row
void printTableRow(String name, String average, String gpa, String grade,
    int nameWidth, int avgWidth, int gpaWidth, int gradeWidth,
    {bool isHeader = false}) {
  final buffer = StringBuffer();
  buffer.write('| ');
  buffer.write(name.padRight(nameWidth - 1));
  buffer.write('| ');
  buffer.write(average.padLeft(avgWidth - 2).padRight(avgWidth - 1));
  buffer.write('| ');
  buffer.write(gpa.padLeft(gpaWidth - 2).padRight(gpaWidth - 1));
  buffer.write('| ');
  buffer.write(grade.padLeft(gradeWidth - 2).padRight(gradeWidth - 1));
  buffer.write('|');
  
  final text = buffer.toString();
  if (isHeader) {
    print('$bold$text$reset');
  } else {
    print(text);
  }
}

/// Prints a table row with color-coded GPA and Grade
void printTableRowWithColors(
    String name, String average, String coloredGpa, String coloredGrade,
    int nameWidth, int avgWidth, int gpaWidth, int gradeWidth) {
  // Calculate the visible length of colored strings (without ANSI codes)
  final gpaVisibleLength = GpaCalculator.formatGpa(
      double.tryParse(coloredGpa.replaceAll(RegExp(r'\x1B\[[0-9;]*m'), '')) ?? 0).length;
  final gradeVisibleLength = coloredGrade.replaceAll(RegExp(r'\x1B\[[0-9;]*m'), '').length;
  
  final buffer = StringBuffer();
  buffer.write('| ');
  buffer.write(name.padRight(nameWidth - 1));
  buffer.write('| ');
  buffer.write(average.padLeft(avgWidth - 2).padRight(avgWidth - 1));
  buffer.write('| ');
  // GPA with color - need to adjust padding for ANSI codes
  final gpaPadding = gpaWidth - 2 - gpaVisibleLength;
  buffer.write(' ' * gpaPadding);
  buffer.write(coloredGpa);
  buffer.write(' ');
  buffer.write('| ');
  // Grade with color - need to adjust padding for ANSI codes
  final gradePadding = gradeWidth - 2 - gradeVisibleLength;
  buffer.write(' ' * gradePadding);
  buffer.write(coloredGrade);
  buffer.write(' ');
  buffer.write('$reset|');
  
  print(buffer.toString());
}

/// Prints summary statistics
void printSummary(List<StudentResult> results) {
  final gpas = results.map((r) => r.gpa).toList();
  final avgGpa = gpas.reduce((a, b) => a + b) / gpas.length;
  final maxGpa = gpas.reduce((a, b) => a > b ? a : b);
  final minGpa = gpas.reduce((a, b) => a < b ? a : b);

  print('${bold}Summary Statistics:$reset');
  print('  Class Average GPA: ${GpaCalculator.getColoredGpa(avgGpa)}');
  print('  Highest GPA: ${GpaCalculator.getColoredGpa(maxGpa)}');
  print('  Lowest GPA: ${GpaCalculator.getColoredGpa(minGpa)}');
  print('');
  
  // Grade distribution
  print('${bold}Grade Distribution:$reset');
  final distribution = <String, int>{};
  for (final result in results) {
    distribution[result.grade] = (distribution[result.grade] ?? 0) + 1;
  }
  
  final grades = ['A', 'B', 'C+', 'C', 'D', 'F'];
  for (final grade in grades) {
    final count = distribution[grade] ?? 0;
    if (count > 0) {
      final bar = '█' * count;
      print('  ${GpaCalculator.getColoredGrade(grade)}: $bar ($count)');
    }
  }
}

/// Prints an error message
void printError(String message) {
  print('${red}Error: $message$reset');
}

/// Prints a success message
void printSuccess(String message) {
  print('${green}✓ $message$reset');
}