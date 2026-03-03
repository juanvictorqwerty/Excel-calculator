package com.example.excelcalculator

import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.InputStream

object ExcelParser {
    
    fun parseExcelFile(inputStream: InputStream): Pair<List<Student>, List<String>>? {
        try {
            val workbook = WorkbookFactory.create(inputStream)
            
            // Find first sheet with "Name" column
            for (sheetIndex in 0 until workbook.numberOfSheets) {
                val sheet = workbook.getSheetAt(sheetIndex)
                if (sheet.physicalNumberOfRows < 2) continue
                
                val headerRow = sheet.getRow(0)
                if (headerRow == null) continue
                
                // Find Name column index
                var nameColumnIndex = -1
                val headers = mutableListOf<String>()
                
                for (cellIndex in 0 until headerRow.physicalNumberOfCells) {
                    val cell = headerRow.getCell(cellIndex)
                    val headerValue = cell?.toString()?.trim() ?: ""
                    headers.add(headerValue)
                    
                    val lowerHeader = headerValue.lowercase()
                    if (lowerHeader == "name" || lowerHeader == "student" || lowerHeader == "student name") {
                        nameColumnIndex = cellIndex
                    }
                }
                
                if (nameColumnIndex == -1) continue
                
                // All other columns are subjects
                val subjectColumnIndices = mutableListOf<Int>()
                val subjectNames = mutableListOf<String>()
                
                for (i in headers.indices) {
                    if (i != nameColumnIndex && headers[i].isNotEmpty()) {
                        subjectColumnIndices.add(i)
                        subjectNames.add(headers[i])
                    }
                }
                
                if (subjectColumnIndices.isEmpty()) continue
                
                // Parse student data
                val students = mutableListOf<Student>()
                
                for (rowIndex in 1 until sheet.physicalNumberOfRows) {
                    val row = sheet.getRow(rowIndex)
                    if (row == null) continue
                    
                    val nameCell = row.getCell(nameColumnIndex)
                    val name = nameCell?.toString()?.trim() ?: ""
                    
                    if (name.isEmpty()) continue
                    
                    val subjects = mutableMapOf<String, Double>()
                    val grades = mutableListOf<Double>()
                    
                    for (colIndex in subjectColumnIndices) {
                        val gradeCell = row.getCell(colIndex)
                        val gradeValue = gradeCell?.toString()?.trim() ?: ""
                        val grade = gradeValue.toDoubleOrNull()
                        
                        if (grade != null) {
                            subjects[headers[colIndex]] = grade
                            grades.add(grade)
                        }
                    }
                    
                    val average = if (grades.isNotEmpty()) {
                        grades.average()
                    } else {
                        0.0
                    }
                    
                    val gpa = GpaCalculator.calculateGpa(average)
                    
                    students.add(Student(name, subjects, average, gpa))
                }
                
                workbook.close()
                return Pair(students, subjectNames)
            }
            
            workbook.close()
            return null
            
        } catch (e: Exception) {
            e.printStackTrace()
            return null
        }
    }
}
