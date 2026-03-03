package com.example.excelcalculator

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.TextView
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.content.ContextCompat
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import com.example.excelcalculator.databinding.ActivityResultsBinding
import com.google.android.material.card.MaterialCardView
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.OutputStream
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

class ResultsActivity : AppCompatActivity() {

    private lateinit var binding: ActivityResultsBinding
    private lateinit var students: ArrayList<Student>
    private lateinit var subjects: ArrayList<String>
    private lateinit var adapter: StudentAdapter

    // ActivityResultLauncher for Storage Access Framework (CREATE_DOCUMENT)
    private val createDocumentLauncher = registerForActivityResult(
        ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ) { uri: Uri? ->
        uri?.let {
            saveExcelToUri(it)
        } ?: run {
            Toast.makeText(this, "Save cancelled", Toast.LENGTH_SHORT).show()
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityResultsBinding.inflate(layoutInflater)
        setContentView(binding.root)

        students = intent.getParcelableArrayListExtra("students") ?: arrayListOf()
        subjects = intent.getStringArrayListExtra("subjects") ?: arrayListOf()

        setupSummary()
        setupRecyclerView()
        setupGpaLegend()

        binding.btnSave.setOnClickListener {
            launchSaveDialog()
        }
    }

    private fun setupSummary() {
        binding.tvTotalStudents.text = "Total Students: ${students.size}"
        val classAverage = if (students.isNotEmpty()) {
            students.map { it.average }.average()
        } else 0.0
        binding.tvClassAverage.text = "Class Average: %.2f".format(classAverage)
    }

    private fun setupRecyclerView() {
        adapter = StudentAdapter(students)
        binding.recyclerView.layoutManager = LinearLayoutManager(this)
        binding.recyclerView.adapter = adapter
    }

    private fun setupGpaLegend() {
        val gpaGrades = listOf(
            "A (80-100)" to GpaCalculator.getGpaColor("A"),
            "B (70-80)" to GpaCalculator.getGpaColor("B"),
            "C+ (60-70)" to GpaCalculator.getGpaColor("C+"),
            "C (50-60)" to GpaCalculator.getGpaColor("C"),
            "D (35-50)" to GpaCalculator.getGpaColor("D"),
            "F (0-35)" to GpaCalculator.getGpaColor("F")
        )

        // Legend is shown when card is expanded
        binding.cardSummary.setOnClickListener {
            val isVisible = binding.layoutLegend.visibility == View.VISIBLE
            binding.layoutLegend.visibility = if (isVisible) View.GONE else View.VISIBLE
        }
    }

    private fun launchSaveDialog() {
        // Generate suggested filename with timestamp to avoid conflicts
        val timestamp = SimpleDateFormat("yyyyMMdd_HHmmss", Locale.getDefault()).format(Date())
        val suggestedFilename = "GPA_Results_$timestamp.xlsx"
        
        // Launch the file picker with suggested filename
        createDocumentLauncher.launch(suggestedFilename)
    }

    private fun saveExcelToUri(uri: Uri) {
        try {
            val workbook = XSSFWorkbook()
            val sheet = workbook.createSheet("GPA Results")

            // Headers
            val headerRow = sheet.createRow(0)
            headerRow.createCell(0).setCellValue("Name")
            headerRow.createCell(1).setCellValue("GPA")

            // Data rows
            students.forEachIndexed { index, student ->
                val row = sheet.createRow(index + 1)
                row.createCell(0).setCellValue(student.name)
                row.createCell(1).setCellValue(student.gpa)
            }

            // Write to the content URI
            contentResolver.openOutputStream(uri)?.use { outputStream: OutputStream ->
                workbook.write(outputStream)
            }
            workbook.close()

            // Get filename from URI for display
            val filename = getFileNameFromUri(uri) ?: "GPA_Results.xlsx"
            Toast.makeText(this, "Saved: $filename", Toast.LENGTH_LONG).show()

        } catch (e: Exception) {
            Toast.makeText(this, "Error saving file: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }

    private fun getFileNameFromUri(uri: Uri): String? {
        // Try to get display name from content resolver
        contentResolver.query(uri, null, null, null, null)?.use { cursor ->
            if (cursor.moveToFirst()) {
                val displayNameIndex = cursor.getColumnIndex(android.provider.OpenableColumns.DISPLAY_NAME)
                if (displayNameIndex != -1) {
                    return cursor.getString(displayNameIndex)
                }
            }
        }
        // Fallback to last path segment
        return uri.lastPathSegment
    }

    inner class StudentAdapter(private val students: List<Student>) :
        RecyclerView.Adapter<StudentAdapter.StudentViewHolder>() {

        inner class StudentViewHolder(itemView: View) : RecyclerView.ViewHolder(itemView) {
            val tvName: TextView = itemView.findViewById(R.id.tvName)
            val tvGpa: TextView = itemView.findViewById(R.id.tvGpa)
            val cardView: MaterialCardView = itemView.findViewById(R.id.cardStudent)
        }

        override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): StudentViewHolder {
            val view = LayoutInflater.from(parent.context)
                .inflate(R.layout.item_student, parent, false)
            return StudentViewHolder(view)
        }

        override fun onBindViewHolder(holder: StudentViewHolder, position: Int) {
            val student = students[position]
            holder.tvName.text = student.name
            holder.tvGpa.text = student.gpa
            
            val gpaColor = GpaCalculator.getGpaColor(student.gpa)
            holder.tvGpa.setBackgroundColor(gpaColor)
            holder.tvGpa.setTextColor(ContextCompat.getColor(this@ResultsActivity, android.R.color.white))
        }

        override fun getItemCount() = students.size
    }
}
