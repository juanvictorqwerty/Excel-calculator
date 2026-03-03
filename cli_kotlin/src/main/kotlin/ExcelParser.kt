import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

/**
 * Exception for Excel parsing errors
 */
class ExcelParseException(message: String) : Exception(message)

/**
 * Handles reading and writing Excel files using Apache POI
 */
object ExcelParser {

    /**
     * Reads an Excel file and extracts student data
     * @param filePath Path to the Excel file
     * @return List of StudentResult objects
     * @throws ExcelParseException if parsing fails
     */
    fun readExcelFile(filePath: String): List<StudentResult> {
        val file = File(filePath)
        
        if (!file.exists()) {
            throw ExcelParseException("File not found: $filePath")
        }

        if (!file.canRead()) {
            throw ExcelParseException("Cannot read file: $filePath")
        }

        FileInputStream(file).use { fis ->
            val workbook = try {
                WorkbookFactory.create(fis)
            } catch (e: Exception) {
                throw ExcelParseException("Failed to parse Excel file: ${e.message}")
            }

            try {
                // Find the sheet with "Name" column
                val sheet = findSheetWithNameColumn(workbook)
                    ?: throw ExcelParseException("No sheet with 'Name' column found")

                return parseSheet(sheet)
            } finally {
                workbook.close()
            }
        }
    }

    /**
     * Finds a sheet that contains a "Name" column in its header row
     * @param workbook The Excel workbook
     * @return The first sheet with a "Name" column, or null if none found
     */
    private fun findSheetWithNameColumn(workbook: Workbook): Sheet? {
        for (i in 0 until workbook.numberOfSheets) {
            val sheet = workbook.getSheetAt(i)
            
            // Check first row for "Name" column
            val firstRow = sheet.getRow(0) ?: continue
            
            for (cell in firstRow) {
                val value = getCellValue(cell)?.toString()?.trim()?.lowercase()
                if (value == "name") {
                    return sheet
                }
            }
        }
        return null
    }

    /**
     * Parses a sheet and extracts student results
     * @param sheet The Excel sheet to parse
     * @return List of StudentResult objects
     */
    private fun parseSheet(sheet: Sheet): List<StudentResult> {
        val results = mutableListOf<StudentResult>()
        
        if (sheet.physicalNumberOfRows < 2) {
            throw ExcelParseException("Sheet must have at least a header row and one data row")
        }

        val headerRow = sheet.getRow(0)
            ?: throw ExcelParseException("Header row is missing")

        // Find the Name column index and subject column indices
        var nameColumnIndex: Int? = null
        val subjectColumnIndices = mutableListOf<Int>()
        
        headerRow.forEachIndexed { index, cell ->
            val value = getCellValue(cell)?.toString()?.trim() ?: ""
            
            when {
                value.isEmpty() -> { /* Skip empty headers */ }
                value.equals("name", ignoreCase = true) -> nameColumnIndex = index
                else -> subjectColumnIndices.add(index)
            }
        }

        if (nameColumnIndex == null) {
            throw ExcelParseException("Name column not found in header")
        }

        if (subjectColumnIndices.isEmpty()) {
            throw ExcelParseException("No subject columns found")
        }

        // Parse data rows
        for (rowIndex in 1..sheet.lastRowNum) {
            val row = sheet.getRow(rowIndex) ?: continue
            
            // Get student name
            val nameCell = row.getCell(nameColumnIndex!!)
            val name = getCellValue(nameCell)?.toString()?.trim() ?: ""
            
            if (name.isEmpty()) continue // Skip empty rows

            // Collect scores from subject columns
            val scores = mutableListOf<Double>()
            for (colIndex in subjectColumnIndices) {
                val cell = row.getCell(colIndex)
                val score = parseScore(getCellValue(cell))
                if (score != null) {
                    scores.add(score)
                }
            }

            if (scores.isNotEmpty()) {
                val average = GpaCalculator.calculateAverage(scores)
                val gpa = GpaCalculator.calculateGpa(scores)
                val grade = GpaCalculator.scoreToGrade(average)
                
                results.add(StudentResult(
                    name = name,
                    average = average,
                    gpa = gpa,
                    grade = grade
                ))
            }
        }

        return results
    }

    /**
     * Extracts the value from a cell based on its type
     * @param cell The Excel cell
     * @return The cell value as Any, or null if empty
     */
    private fun getCellValue(cell: Cell?): Any? {
        if (cell == null) return null

        return when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue
            CellType.NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    cell.dateCellValue
                } else {
                    cell.numericCellValue
                }
            }
            CellType.BOOLEAN -> cell.booleanCellValue
            CellType.FORMULA -> {
                // Try to evaluate the formula
                try {
                    val evaluator = cell.sheet.workbook.creationHelper.createFormulaEvaluator()
                    val evaluatedCell = evaluator.evaluate(cell)
                    when (evaluatedCell.cellType) {
                        CellType.STRING -> evaluatedCell.stringValue
                        CellType.NUMERIC -> evaluatedCell.numberValue
                        CellType.BOOLEAN -> evaluatedCell.booleanValue
                        else -> null
                    }
                } catch (e: Exception) {
                    null
                }
            }
            CellType.BLANK -> null
            CellType.ERROR -> null
            else -> null
        }
    }

    /**
     * Parses a cell value into a score (double)
     * @param value The cell value
     * @return The parsed score, or null if invalid
     */
    private fun parseScore(value: Any?): Double? {
        return when (value) {
            is Number -> value.toDouble()
            is String -> {
                val cleaned = value.trim()
                if (cleaned.isEmpty()) return null
                cleaned.toDoubleOrNull()
            }
            else -> null
        }
    }

    /**
     * Saves student results to a new Excel file with timestamp
     * @param results List of StudentResult objects
     * @param inputFilePath Path to the original input file (for directory reference)
     * @return Path to the saved file
     * @throws ExcelParseException if saving fails
     */
    fun saveResultsToExcel(results: List<StudentResult>, inputFilePath: String): String {
        val timestamp = LocalDateTime.now().format(
            DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss")
        )
        
        val inputFile = File(inputFilePath)
        val directory = inputFile.parentFile?.absolutePath ?: "."
        val outputFileName = "GPA_Results_$timestamp.xlsx"
        val outputPath = "$directory/$outputFileName"

        XSSFWorkbook().use { workbook ->
            val sheet = workbook.createSheet("GPA Results")
            
            // Create header row
            val headerRow = sheet.createRow(0)
            val headers = listOf("Name", "Average", "GPA", "Grade")
            headers.forEachIndexed { index, header ->
                val cell = headerRow.createCell(index)
                cell.setCellValue(header)
                
                // Style the header
                val style = workbook.createCellStyle()
                val font = workbook.createFont()
                font.bold = true
                style.setFont(font)
                cell.cellStyle = style
            }

            // Add data rows
            results.forEachIndexed { index, result ->
                val row = sheet.createRow(index + 1)
                row.createCell(0).setCellValue(result.name)
                row.createCell(1).setCellValue(result.average)
                row.createCell(2).setCellValue(result.gpa)
                row.createCell(3).setCellValue(result.grade)
            }

            // Auto-size columns
            for (i in 0..3) {
                sheet.autoSizeColumn(i)
            }

            // Save the file
            FileOutputStream(outputPath).use { fos ->
                workbook.write(fos)
            }
        }

        return outputPath
    }

    /**
     * Gets the subject column names from the Excel file
     * @param filePath Path to the Excel file
     * @return List of subject column names
     * @throws ExcelParseException if parsing fails
     */
    fun getSubjectNames(filePath: String): List<String> {
        val file = File(filePath)
        
        if (!file.exists()) {
            throw ExcelParseException("File not found: $filePath")
        }

        FileInputStream(file).use { fis ->
            val workbook = WorkbookFactory.create(fis)
            
            try {
                val sheet = findSheetWithNameColumn(workbook)
                    ?: throw ExcelParseException("No sheet with 'Name' column found")

                val headerRow = sheet.getRow(0) 
                    ?: return emptyList()

                val subjects = mutableListOf<String>()
                headerRow.forEach { cell ->
                    val value = getCellValue(cell)?.toString()?.trim() ?: ""
                    if (value.isNotEmpty() && !value.equals("name", ignoreCase = true)) {
                        subjects.add(value)
                    }
                }

                return subjects
            } finally {
                workbook.close()
            }
        }
    }
}
