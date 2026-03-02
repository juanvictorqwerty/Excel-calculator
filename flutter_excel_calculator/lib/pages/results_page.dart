import 'dart:io';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;
import 'package:path_provider/path_provider.dart';
import '../models/student.dart';

class ResultsPage extends StatelessWidget {
  final List<Student> students;
  final List<String> subjectColumns;

  const ResultsPage({
    super.key,
    required this.students,
    required this.subjectColumns,
  });

  Color _getGPAColor(String gpa) {
    switch (gpa) {
      case 'A': return Colors.green.shade700;
      case 'B': return Colors.blue.shade700;
      case 'C+': return Colors.teal.shade700;
      case 'C': return Colors.orange.shade700;
      case 'D': return Colors.deepOrange.shade700;
      case 'F': return Colors.red.shade700;
      default: return Colors.grey;
    }
  }

  /// Save results as a new Excel file
  Future<void> _saveResultsAsExcel(BuildContext context) async {
    try {
      final workbook = xlsio.Workbook();
      final sheet = workbook.worksheets[0];
      sheet.name = 'GPA Results';

      // Headers
      sheet.getRangeByIndex(1, 1).setText('Student Name');
      sheet.getRangeByIndex(1, 2).setText('Average');
      sheet.getRangeByIndex(1, 3).setText('GPA');
      
      // Add subject columns
      for (int i = 0; i < subjectColumns.length; i++) {
        sheet.getRangeByIndex(1, 4 + i).setText(subjectColumns[i]);
      }

      // Style header row
      for (int i = 1; i <= 3 + subjectColumns.length; i++) {
        final cell = sheet.getRangeByIndex(1, i);
        cell.cellStyle.bold = true;
        cell.cellStyle.backColor = '#CCCCCC';
      }

      // Data rows
      for (int i = 0; i < students.length; i++) {
        final student = students[i];
        final rowIndex = i + 2;
        
        sheet.getRangeByIndex(rowIndex, 1).setText(student.name);
        sheet.getRangeByIndex(rowIndex, 2).setNumber(student.average);
        sheet.getRangeByIndex(rowIndex, 3).setText(student.gpa);
        
        // Add subject grades
        for (int j = 0; j < subjectColumns.length; j++) {
          final subject = subjectColumns[j];
          final grade = student.subjects[subject];
          if (grade != null) {
            sheet.getRangeByIndex(rowIndex, 4 + j).setNumber(grade);
          }
        }
      }

      // Save file
      final bytes = workbook.saveAsStream();
      workbook.dispose();

      // Get save location
      String? outputFile;
      if (Platform.isAndroid) {
        // Save to Downloads folder on Android
        outputFile = '/storage/emulated/0/Download/GPA_Results.xlsx';
      } else if (Platform.isIOS) {
        final directory = await getApplicationDocumentsDirectory();
        outputFile = '${directory.path}/GPA_Results.xlsx';
      } else {
        outputFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Save GPA Results',
          fileName: 'GPA_Results.xlsx',
          allowedExtensions: ['xlsx'],
          type: FileType.custom,
        );
      }

      if (outputFile != null) {
        final file = File(outputFile);
        await file.writeAsBytes(bytes);
        
        if (context.mounted) {
          ScaffoldMessenger.of(context).showSnackBar(
            SnackBar(content: Text('Saved to: $outputFile')),
          );
        }
      } else if (context.mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Save cancelled')),
        );
      }
    } catch (e) {
      if (context.mounted) {
        showDialog(
          context: context,
          builder: (context) => AlertDialog(
            title: const Text('Error'),
            content: Text('Error saving file: $e'),
            actions: [
              TextButton(
                onPressed: () => Navigator.pop(context),
                child: const Text('OK'),
              ),
            ],
          ),
        );
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    final classAverage = students.map((s) => s.average).reduce((a, b) => a + b) / students.length;
    
    return Scaffold(
      appBar: AppBar(
        title: const Text('GPA Results'),
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        actions: [
          IconButton(
            onPressed: () => _saveResultsAsExcel(context),
            icon: const Icon(Icons.save),
            tooltip: 'Save as Excel',
          ),
        ],
      ),
      body: Column(
        children: [
          // Summary Card
          Card(
            margin: const EdgeInsets.all(16),
            child: Padding(
              padding: const EdgeInsets.all(16),
              child: Row(
                mainAxisAlignment: MainAxisAlignment.spaceAround,
                children: [
                  _buildStat('Total Students', students.length.toString()),
                  _buildStat('Class Average', classAverage.toStringAsFixed(2)),
                ],
              ),
            ),
          ),
          
          // GPA Legend
          Padding(
            padding: const EdgeInsets.symmetric(horizontal: 16),
            child: Wrap(
              spacing: 8,
              children: [
                _buildLegend('A (80-100)', Colors.green.shade700),
                _buildLegend('B (70-80)', Colors.blue.shade700),
                _buildLegend('C+ (60-70)', Colors.teal.shade700),
                _buildLegend('C (50-60)', Colors.orange.shade700),
                _buildLegend('D (35-50)', Colors.deepOrange.shade700),
                _buildLegend('F (0-35)', Colors.red.shade700),
              ],
            ),
          ),
          
          const SizedBox(height: 8),
          const Divider(),
          
          // Students List
          Expanded(
            child: ListView.builder(
              itemCount: students.length,
              itemBuilder: (context, index) {
                final student = students[index];
                return Card(
                  margin: const EdgeInsets.symmetric(horizontal: 16, vertical: 4),
                  child: ExpansionTile(
                    title: Row(
                      children: [
                        Expanded(
                          child: Text(
                            student.name,
                            style: const TextStyle(fontWeight: FontWeight.bold),
                          ),
                        ),
                        Container(
                          padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 4),
                          decoration: BoxDecoration(
                            color: _getGPAColor(student.gpa),
                            borderRadius: BorderRadius.circular(12),
                          ),
                          child: Text(
                            student.gpa,
                            style: const TextStyle(
                              color: Colors.white,
                              fontWeight: FontWeight.bold,
                            ),
                          ),
                        ),
                      ],
                    ),
                    subtitle: Text('Average: ${student.average.toStringAsFixed(2)}'),
                    children: [
                      Padding(
                        padding: const EdgeInsets.all(16),
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            const Text(
                              'Subject Grades:',
                              style: TextStyle(fontWeight: FontWeight.bold),
                            ),
                            const SizedBox(height: 8),
                            ...student.subjects.entries.map((entry) {
                              return Padding(
                                padding: const EdgeInsets.symmetric(vertical: 2),
                                child: Row(
                                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                                  children: [
                                    Text(entry.key),
                                    Text(
                                      entry.value.toStringAsFixed(2),
                                      style: const TextStyle(fontWeight: FontWeight.bold),
                                    ),
                                  ],
                                ),
                              );
                            }),
                          ],
                        ),
                      ),
                    ],
                  ),
                );
              },
            ),
          ),
          
          // Save Button
          Padding(
            padding: const EdgeInsets.all(16),
            child: ElevatedButton.icon(
              onPressed: () => _saveResultsAsExcel(context),
              icon: const Icon(Icons.save),
              label: const Text('Save Results as Excel'),
              style: ElevatedButton.styleFrom(
                padding: const EdgeInsets.symmetric(vertical: 16, horizontal: 32),
              ),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildStat(String label, String value) {
    return Column(
      children: [
        Text(
          value,
          style: const TextStyle(
            fontSize: 24,
            fontWeight: FontWeight.bold,
          ),
        ),
        Text(label),
      ],
    );
  }

  Widget _buildLegend(String label, Color color) {
    return Chip(
      avatar: CircleAvatar(
        backgroundColor: color,
        radius: 8,
      ),
      label: Text(label, style: const TextStyle(fontSize: 10)),
      backgroundColor: Colors.grey.shade100,
    );
  }
}
