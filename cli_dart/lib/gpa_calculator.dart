/// GPA Calculator - Contains logic for calculating GPA from scores
class GpaCalculator {
  /// Converts a numeric score to a letter grade
  static String scoreToGrade(double score) {
    if (score <= 50) return 'D';
    if (score <= 60) return 'C';
    if (score <= 70) return 'C+';
    if (score <= 80) return 'B';
    if (score >= 80) return 'A';
    
    return 'F';
  }

  /// Converts a letter grade to grade points
  static double gradeToPoints(String grade) {
    switch (grade) {
      case 'A':
        return 4.0;
      case 'B':
        return 3.0;
      case 'C+':
        return 2.5;
      case 'C':
        return 2.0;
      case 'D':
        return 1.0;
      case 'F':
        return 0.0;
      default:
        return 0.0;
    }
  }

  /// Gets the ANSI color code for a grade
  static String getGradeColor(String grade) {
    // ANSI color codes
    const reset = '\x1B[0m';
    switch (grade) {
      case 'A':
        return '\x1B[32m'; // Green
      case 'B':
        return '\x1B[34m'; // Blue
      case 'C+':
        return '\x1B[36m'; // Cyan
      case 'C':
        return '\x1B[33m'; // Yellow
      case 'D':
        return '\x1B[38;5;208m'; // Orange (256-color mode)
      case 'F':
        return '\x1B[31m'; // Red
      default:
        return '\x1B[37m'; // White
    }
  }

  /// Calculates the average score from a list of scores
  static double calculateAverage(List<double> scores) {
    if (scores.isEmpty) return 0.0;
    final sum = scores.reduce((a, b) => a + b);
    return sum / scores.length;
  }

  /// Calculates GPA from a list of scores
  static double calculateGpa(List<double> scores) {
    if (scores.isEmpty) return 0.0;
    
    double totalPoints = 0.0;
    for (final score in scores) {
      final grade = scoreToGrade(score);
      totalPoints += gradeToPoints(grade);
    }
    
    return totalPoints / scores.length;
  }

  /// Formats GPA to 2 decimal places
  static String formatGpa(double gpa) {
    return gpa.toStringAsFixed(2);
  }

  /// Gets color-coded GPA string for console output
  static String getColoredGpa(double gpa) {
    final grade = scoreToGrade(gpa * 25); // Scale GPA (0-4) to score (0-100) approximately
    final color = getGradeColor(grade);
    final reset = '\x1B[0m';
    return '$color${formatGpa(gpa)}$reset';
  }

  /// Gets color-coded grade string for console output
  static String getColoredGrade(String grade) {
    final color = getGradeColor(grade);
    final reset = '\x1B[0m';
    return '$color$grade$reset';
  }
}

/// Represents a student's GPA result
class StudentResult {
  final String name;
  final double average;
  final double gpa;
  final String grade;

  StudentResult({
    required this.name,
    required this.average,
    required this.gpa,
    required this.grade,
  });

  @override
  String toString() {
    return 'StudentResult(name: $name, average: $average, gpa: $gpa, grade: $grade)';
  }
}