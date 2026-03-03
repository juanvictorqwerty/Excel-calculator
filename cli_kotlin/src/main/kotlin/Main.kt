import java.io.File
import java.io.IOException
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

/**
 * Main entry point for the GPA Calculator CLI application
 */
fun main(args: Array<String>) {
    println()
    printHeader()

    // Get file path from command line argument or prompt user
    val filePath = when {
        args.isNotEmpty() -> args[0]
        else -> promptForFilePath()
    }

    if (filePath.isBlank()) {
        printError("No file path provided. Exiting.")
        return
    }

    // Validate file exists
    val file = File(filePath)
    if (!file.exists()) {
        printError("File not found: $filePath")
        println("Please check the path and try again.")
        return
    }

    // Validate file extension
    val extension = file.extension.lowercase()
    if (extension != "xlsx" && extension != "xls") {
        printError("Invalid file format. Expected .xlsx or .xls, got .$extension")
        return
    }

    try {
        // Parse Excel file
        println("\n${AnsiColors.CYAN}Reading Excel file: $filePath${AnsiColors.RESET}")
        val results = ExcelParser.readExcelFile(filePath)

        if (results.isEmpty()) {
            printError("No valid student data found in the Excel file.")
            return
        }

        // Display results
        println("\n${AnsiColors.BOLD}${AnsiColors.GREEN}✓ Successfully parsed ${results.size} student(s)${AnsiColors.RESET}\n")
        displayResults(results)

        // Save results to new Excel file
        val outputPath = ExcelParser.saveResultsToExcel(results, filePath)
        println("\n${AnsiColors.BOLD}${AnsiColors.GREEN}✓ Results saved to: $outputPath${AnsiColors.RESET}\n")

    } catch (e: ExcelParseException) {
        printError("Excel parsing error: ${e.message}")
    } catch (e: IOException) {
        printError("IO error: ${e.message}")
    } catch (e: Exception) {
        printError("Unexpected error: ${e.message}")
        e.printStackTrace()
    }
}

/**
 * Prints the application header
 */
private fun printHeader() {
    println("${AnsiColors.BOLD}${AnsiColors.CYAN}")
    println("╔══════════════════════════════════════════════════════════════╗")
    println("║                GPA Calculator - Console Edition              ║")
    println("║                        Version 1.0.0                         ║")
    println("╚══════════════════════════════════════════════════════════════╝")
    println("${AnsiColors.RESET}")
}

/**
 * Prompts the user for a file path
 */
private fun promptForFilePath(): String {
    println("${AnsiColors.YELLOW}Please enter the path to the Excel file:${AnsiColors.RESET}")
    print("> ")
    return readlnOrNull()?.trim() ?: ""
}

/**
 * Prints an error message with color
 */
private fun printError(message: String) {
    println("${AnsiColors.BOLD}${AnsiColors.RED}✗ Error: $message${AnsiColors.RESET}")
}

/**
 * Displays the results in a formatted table
 */
private fun displayResults(results: List<StudentResult>) {
    // Calculate column widths
    val maxNameLength = maxOf(
        "Name".length,
        results.maxOfOrNull { it.name.length } ?: 0
    )
    val nameWidth = maxNameLength + 2

    // Print table header
    println("${AnsiColors.BOLD}${AnsiColors.WHITE}")
    println("┌${"─".repeat(nameWidth)}┬─────────────┬─────────┬───────┐")
    println("│${" Name".padEnd(nameWidth)}│ Average     │ GPA     │ Grade │")
    println("├${"─".repeat(nameWidth)}┼─────────────┼─────────┼───────┤")
    println("${AnsiColors.RESET}")

    // Print student rows
    results.forEach { result ->
        val coloredGpa = GpaCalculator.getColoredGpa(result.gpa)
        val coloredGrade = GpaCalculator.getColoredGrade(result.grade)
        val average = String.format("%.2f", result.average).padEnd(11)
        
        println(
            "│ ${AnsiColors.WHITE}${result.name.padEnd(nameWidth - 1)}${AnsiColors.RESET}│ " +
            "${AnsiColors.WHITE}$average${AnsiColors.RESET}│ " +
            "${coloredGpa.padEnd(17)}${AnsiColors.RESET}│ " +
            "${coloredGrade.padEnd(15)}${AnsiColors.RESET}│"
        )
    }

    // Print table footer
    println("${AnsiColors.BOLD}${AnsiColors.WHITE}")
    println("└${"─".repeat(nameWidth)}┴─────────────┴─────────┴───────┘")
    println("${AnsiColors.RESET}")

    // Print summary statistics
    printSummary(results)
}

/**
 * Prints summary statistics
 */
private fun printSummary(results: List<StudentResult>) {
    val avgGpa = results.map { it.gpa }.average()
    val avgScore = results.map { it.average }.average()
    val highestGpa = results.maxByOrNull { it.gpa }
    val lowestGpa = results.minByOrNull { it.gpa }

    println("\n${AnsiColors.BOLD}${AnsiColors.CYAN}Summary Statistics:${AnsiColors.RESET}")
    println("  • Total Students: ${results.size}")
    println("  • Class Average Score: ${AnsiColors.YELLOW}${String.format("%.2f", avgScore)}${AnsiColors.RESET}")
    println("  • Class Average GPA: ${GpaCalculator.getColoredGpa(avgGpa)}")
    
    highestGpa?.let {
        println("  • Highest GPA: ${GpaCalculator.getColoredGpa(it.gpa)} (${it.name})")
    }
    lowestGpa?.let {
        println("  • Lowest GPA: ${GpaCalculator.getColoredGpa(it.gpa)} (${it.name})")
    }
}
