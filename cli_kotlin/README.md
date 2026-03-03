# GPA Calculator - Kotlin Console Edition

A command-line application for calculating GPA from Excel files containing student grades.

## Features

- Reads Excel files (.xlsx, .xls) with student names and subject grades
- Automatically finds the sheet with "Name" column
- Calculates GPA and letter grades for all students
- Color-coded console output:
  - A = Green
  - B = Blue
  - C+ = Cyan
  - C = Yellow
  - D = Orange
  - F = Red
- Saves results to a new Excel file with timestamp
- Handles errors gracefully

## Requirements

- Java 21 or higher
- Gradle 8.7 (included via wrapper)

## Building

### Using Gradle Wrapper (Recommended)

```bash
./gradlew build
```

### Creating Fat JAR (Standalone)

```bash
./gradlew jar
```

The JAR will be created at `build/libs/cli_kotlin-1.0.0.jar`

## Usage

### Command Line with Argument

```bash
# Using the fat JAR
java -jar build/libs/cli_kotlin-1.0.0.jar path/to/grades.xlsx

# Using Gradle
./gradlew run --args="path/to/grades.xlsx"
```

### Interactive Mode (Prompt for Path)

```bash
java -jar build/libs/cli_kotlin-1.0.0.jar
```

Then enter the file path when prompted.

## Excel File Format

Your Excel file should have the following format:

| Name          | Math | Science | English | History |
| ------------- | ---- | ------- | ------- | ------- |
| Alice Johnson | 95   | 88      | 92      | 85      |
| Bob Smith     | 78   | 82      | 75      | 80      |

- The first row contains headers
- One column must be named "Name" (case-insensitive)
- All other columns are treated as subjects with numeric scores
- Supports scores from 0-100

## Output

The application displays:

- Formatted table with Name, Average, GPA, and Grade
- Summary statistics (Class average, highest/lowest GPA)
- Results are saved to `GPA_Results_[timestamp].xlsx` in the same directory as the input file

## Example Output

```
╔══════════════════════════════════════════════════════════════╗
║                GPA Calculator - Console Edition              ║
║                        Version 1.0.0                         ║
╚══════════════════════════════════════════════════════════════╝

✓ Successfully parsed 8 student(s)

┌───────────────┬─────────────┬─────────┬───────┐
│ Name          │ Average     │ GPA     │ Grade │
├───────────────┼─────────────┼─────────┼───────┤
│ Alice Johnson │ 89.40       │ 3.60    │ B     │
│ Bob Smith     │ 78.60       │ 2.70    │ C+    │
│ Charlie Brown │ 68.40       │ 2.20    │ C     │
└───────────────┴─────────────┴─────────┴───────┘

Summary Statistics:
  • Total Students: 8
  • Class Average Score: 74.08
  • Class Average GPA: 2.45

✓ Results saved to: ./GPA_Results_20260303_104226.xlsx
```

## Error Handling

The application handles the following errors gracefully:

- File not found
- Invalid file format (non-Excel files)
- Missing "Name" column
- No subject columns found
- Empty sheets
- Invalid numeric values

## Project Structure

```
cli_kotlin/
├── build.gradle.kts          # Gradle build configuration
├── gradlew                   # Gradle wrapper (Linux/Mac)
├── gradlew.bat               # Gradle wrapper (Windows)
├── gradle/wrapper/           # Gradle wrapper files
├── README.md                 # This file
├── sample_grades.xlsx        # Sample data file
└── src/main/kotlin/
    ├── Main.kt               # Application entry point
    ├── GpaCalculator.kt      # GPA calculation logic
    └── ExcelParser.kt        # Excel reading/writing
```

## Technologies Used

- Kotlin 1.9.22
- Apache POI 5.2.5 (for Excel handling)
- Gradle 8.7

## License

MIT License
