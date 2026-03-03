/**
 * ANSI color codes for terminal output
 */
object AnsiColors {
    const val RESET = "\u001B[0m"
    const val BOLD = "\u001B[1m"
    const val BLACK = "\u001B[30m"
    const val RED = "\u001B[31m"
    const val GREEN = "\u001B[32m"
    const val YELLOW = "\u001B[33m"
    const val BLUE = "\u001B[34m"
    const val MAGENTA = "\u001B[35m"
    const val CYAN = "\u001B[36m"
    const val WHITE = "\u001B[37m"
    const val ORANGE = "\u001B[38;5;208m" // 256-color mode for orange
}

/**
 * Represents a student's GPA result
 */
data class StudentResult(
    val name: String,
    val average: Double,
    val gpa: Double,
    val grade: String
)

/**
 * GPA Calculator - Contains logic for calculating GPA from scores
 */
object GpaCalculator {

    /**
     * Converts a numeric score to a letter grade
     * @param score The numeric score (0-100)
     * @return The letter grade (A, B, C+, C, D, F)
     */
    fun scoreToGrade(score: Double): String {
        return when {
            score <= 50 -> "D"
            score <= 60 -> "C"
            score <= 70 -> "C+"
            score <= 80 -> "B"
            score >= 80 -> "A"

            else -> "F"
        }
    }

    /**
     * Converts a letter grade to grade points
     * @param grade The letter grade
     * @return The grade points (0.0 - 4.0)
     */
    fun gradeToPoints(grade: String): Double {
        return when (grade) {
            "A" -> 4.0
            "B" -> 3.0
            "C+" -> 2.5
            "C" -> 2.0
            "D" -> 1.0
            "F" -> 0.0
            else -> 0.0
        }
    }

    /**
     * Gets the ANSI color code for a grade
     * @param grade The letter grade
     * @return The ANSI color code string
     */
    fun getGradeColor(grade: String): String {
        return when (grade) {
            "A" -> AnsiColors.GREEN
            "B" -> AnsiColors.BLUE
            "C+" -> AnsiColors.CYAN
            "C" -> AnsiColors.YELLOW
            "D" -> AnsiColors.ORANGE
            "F" -> AnsiColors.RED
            else -> AnsiColors.WHITE
        }
    }

    /**
     * Calculates the average score from a list of scores
     * @param scores List of numeric scores
     * @return The average score
     */
    fun calculateAverage(scores: List<Double>): Double {
        if (scores.isEmpty()) return 0.0
        return scores.average()
    }

    /**
     * Calculates GPA from a list of scores
     * @param scores List of numeric scores
     * @return The calculated GPA (0.0 - 4.0)
     */
    fun calculateGpa(scores: List<Double>): Double {
        if (scores.isEmpty()) return 0.0

        val totalPoints = scores.sumOf { score ->
            val grade = scoreToGrade(score)
            gradeToPoints(grade)
        }

        return totalPoints / scores.size
    }

    /**
     * Formats GPA to 2 decimal places
     * @param gpa The GPA value
     * @return Formatted GPA string
     */
    fun formatGpa(gpa: Double): String {
        return String.format("%.2f", gpa)
    }

    /**
     * Gets color-coded GPA string for console output
     * @param gpa The GPA value
     * @return Color-coded GPA string with ANSI escape codes
     */
    fun getColoredGpa(gpa: Double): String {
        // Scale GPA (0-4) to score (0-100) approximately for grade determination
        val score = gpa * 25
        val grade = scoreToGrade(score)
        val color = getGradeColor(grade)
        return "$color${formatGpa(gpa)}${AnsiColors.RESET}"
    }

    /**
     * Gets color-coded grade string for console output
     * @param grade The letter grade
     * @return Color-coded grade string with ANSI escape codes
     */
    fun getColoredGrade(grade: String): String {
        val color = getGradeColor(grade)
        return "$color$grade${AnsiColors.RESET}"
    }

    /**
     * Gets a color bar visualization for GPA
     * @param gpa The GPA value
     * @param width The width of the bar in characters
     * @return A color-coded progress bar string
     */
    fun getGpaBar(gpa: Double, width: Int = 20): String {
        val score = gpa * 25
        val grade = scoreToGrade(score)
        val color = getGradeColor(grade)
        val filled = ((gpa / 4.0) * width).toInt().coerceIn(0, width)
        val empty = width - filled
        
        return "$color${"█".repeat(filled)}${"░".repeat(empty)}${AnsiColors.RESET}"
    }
}
