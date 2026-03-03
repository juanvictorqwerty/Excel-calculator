package com.example.excelcalculator

object GpaCalculator {
    
    fun calculateGpa(average: Double): String {
        return when {
            average >= 80 -> "A"
            average >= 70 -> "B"
            average >= 60 -> "C+"
            average >= 50 -> "C"
            average >= 35 -> "D"
            else -> "F"
        }
    }
    
    fun getGpaColor(gpa: String): Int {
        return when (gpa) {
            "A" -> 0xFF2E7D32.toInt()      // Green
            "B" -> 0xFF1976D2.toInt()      // Blue
            "C+" -> 0xFF00897B.toInt()     // Teal
            "C" -> 0xFFF57C00.toInt()      // Orange
            "D" -> 0xFFE64A19.toInt()      // Deep Orange
            "F" -> 0xFFD32F2F.toInt()      // Red
            else -> 0xFF757575.toInt()     // Grey
        }
    }
}
