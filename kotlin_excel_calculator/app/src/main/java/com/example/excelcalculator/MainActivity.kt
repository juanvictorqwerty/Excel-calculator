package com.example.excelcalculator

import android.Manifest
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.provider.OpenableColumns
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import com.example.excelcalculator.databinding.ActivityMainBinding

class MainActivity : AppCompatActivity() {

    private lateinit var binding: ActivityMainBinding
    private val REQUEST_CODE_PERMISSIONS = 1001

    private val filePickerLauncher = registerForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri: Uri? ->
        uri?.let {
            handleSelectedFile(it)
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        binding.btnBrowse.setOnClickListener {
            checkPermissionsAndPickFile()
        }
    }

    private fun checkPermissionsAndPickFile() {
        when {
            Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q -> {
                // Android 10+ doesn't need storage permission for file picker
                openFilePicker()
            }
            ContextCompat.checkSelfPermission(
                this,
                Manifest.permission.READ_EXTERNAL_STORAGE
            ) == PackageManager.PERMISSION_GRANTED -> {
                openFilePicker()
            }
            else -> {
                ActivityCompat.requestPermissions(
                    this,
                    arrayOf(Manifest.permission.READ_EXTERNAL_STORAGE),
                    REQUEST_CODE_PERMISSIONS
                )
            }
        }
    }

    private fun openFilePicker() {
        filePickerLauncher.launch(arrayOf(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/vnd.ms-excel"
        ))
    }

    private fun handleSelectedFile(uri: Uri) {
        try {
            contentResolver.openInputStream(uri)?.use { inputStream ->
                val result = ExcelParser.parseExcelFile(inputStream)
                
                if (result != null) {
                    val (students, subjects) = result
                    if (students.isNotEmpty()) {
                        val intent = Intent(this, ResultsActivity::class.java).apply {
                            putParcelableArrayListExtra("students", ArrayList(students))
                            putStringArrayListExtra("subjects", ArrayList(subjects))
                        }
                        startActivity(intent)
                    } else {
                        Toast.makeText(this, "No student data found", Toast.LENGTH_LONG).show()
                    }
                } else {
                    Toast.makeText(
                        this,
                        "No sheet with 'Name' column found. Please check your file format.",
                        Toast.LENGTH_LONG
                    ).show()
                }
            }
        } catch (e: Exception) {
            Toast.makeText(this, "Error reading file: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }

    override fun onRequestPermissionsResult(
        requestCode: Int,
        permissions: Array<out String>,
        grantResults: IntArray
    ) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        if (requestCode == REQUEST_CODE_PERMISSIONS) {
            if (grantResults.isNotEmpty() && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                openFilePicker()
            } else {
                Toast.makeText(this, "Storage permission required", Toast.LENGTH_LONG).show()
            }
        }
    }
}
