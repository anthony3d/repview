package com.example.repview

import android.os.Bundle
import android.widget.TableLayout
import android.widget.TableRow
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.util.*

class ReportViewerActivity : AppCompatActivity() {
    
    private lateinit var tableLayout: TableLayout
    private val dateFormat = SimpleDateFormat("dd-MM-yyyy", Locale.getDefault())
    
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_report_viewer)
        
        tableLayout = findViewById(R.id.tableLayout)
        
        // Проверяем, что нас вызвали через Share
        when {
            intent?.action == Intent.ACTION_SEND -> {
                val uri = intent.getParcelableExtra<android.net.Uri>(Intent.EXTRA_STREAM)
                if (uri != null) {
                    processFile(uri)
                } else {
                    Toast.makeText(this, "Файл не найден", Toast.LENGTH_SHORT).show()
                    finish()
                }
            }
            else -> {
                Toast.makeText(this, "Запустите через 'Поделиться'", Toast.LENGTH_SHORT).show()
                finish()
            }
        }
    }
    
    private fun processFile(uri: android.net.Uri) {
        try {
            // Копируем файл во временный файл
            val inputStream = contentResolver.openInputStream(uri)
            val tempFile = File(cacheDir, "temp_report.xlsx")
            tempFile.outputStream().use { output ->
                inputStream?.copyTo(output)
            }
            inputStream?.close()
            
            // Читаем XLSX
            val workbook = WorkbookFactory.create(tempFile)
            val sheet = workbook.getSheetAt(0) // Первый лист
            
            // Получаем заголовки (первая строка)
            val headerRow = sheet.getRow(0)
            val lastColumnIndex = headerRow.lastCellNum - 1
            val preLastColumnIndex = lastColumnIndex - 1
            
            // Добавляем заголовки нашей таблицы
            addTableHeader()
            
            // Проходим по всем строкам данных (со второй строки)
            for (i in 1 until sheet.physicalNumberOfRows) {
                val row = sheet.getRow(i) ?: continue
                
                // Получаем столбцы (индексация с 0)
                // Столбец 3 = индекс 2
                val startDateCell = row.getCell(2) 
                // Столбец 4 = индекс 3
                val endDateCell = row.getCell(3)
                // Предпоследний столбец
                val numberCell = row.getCell(preLastColumnIndex)
                
                // Извлекаем даты
                val startDate = parseDate(startDateCell)
                val endDate = parseDate(endDateCell)
                val number = if (numberCell != null) numberCell.numericCellValue.toInt().toString() else "0"
                
                if (startDate != null) {
                    val startDateStr = dateFormat.format(startDate)
                    val endDateStr = if (endDate != null) dateFormat.format(endDate) else "-"
                    val plus4Weeks = addWeeks(startDate, 4)
                    val plus5Weeks = addWeeks(startDate, 5)
                    
                    addDataRow(startDateStr, endDateStr, number, 
                               dateFormat.format(plus4Weeks), dateFormat.format(plus5Weeks))
                }
            }
            
            workbook.close()
            tempFile.delete()
            
            if (sheet.physicalNumberOfRows <= 1) {
                Toast.makeText(this, "Нет данных в таблице", Toast.LENGTH_SHORT).show()
            }
            
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Ошибка: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }
    
    private fun parseDate(cell: org.apache.poi.ss.usermodel.Cell?): Date? {
        if (cell == null) return null
        
        return try {
            when (cell.cellType) {
                org.apache.poi.ss.usermodel.CellType.NUMERIC -> {
                    if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                        cell.dateCellValue
                    } else {
                        null
                    }
                }
                org.apache.poi.ss.usermodel.CellType.STRING -> {
                    // Пробуем распарсить строку
                    val formats = listOf(
                        SimpleDateFormat("dd.MM.yyyy", Locale.getDefault()),
                        SimpleDateFormat("dd/MM/yyyy", Locale.getDefault()),
                        SimpleDateFormat("yyyy-MM-dd", Locale.getDefault())
                    )
                    for (format in formats) {
                        try {
                            return format.parse(cell.stringCellValue)
                        } catch (e: Exception) {
                            // Пробуем следующий формат
                        }
                    }
                    null
                }
                else -> null
            }
        } catch (e: Exception) {
            null
        }
    }
    
    private fun addWeeks(date: Date, weeks: Int): Date {
        val calendar = Calendar.getInstance()
        calendar.time = date
        calendar.add(Calendar.WEEK_OF_YEAR, weeks)
        return calendar.time
    }
    
    private fun addTableHeader() {
        val row = TableRow(this)
        val headers = arrayOf("Начало недели", "Конец недели", "Значение", "+4 недели", "+5 недель")
        
        for (header in headers) {
            val tv = TextView(this).apply {
                text = header
                setPadding(16, 12, 16, 12)
                setTextColor(resources.getColor(android.R.color.white))
                setBackgroundColor(resources.getColor(android.R.color.holo_blue_dark))
                textSize = 14f
            }
            row.addView(tv)
        }
        tableLayout.addView(row)
    }
    
    private fun addDataRow(col1: String, col2: String, col3: String, col4: String, col5: String) {
        val row = TableRow(this)
        val data = arrayOf(col1, col2, col3, col4, col5)
        
        for ((index, value) in data.withIndex()) {
            val tv = TextView(this).apply {
                text = value
                setPadding(16, 12, 16, 12)
                textSize = 12f
                if (index % 2 == 0) {
                    setBackgroundColor(resources.getColor(android.R.color.white))
                } else {
                    setBackgroundColor(resources.getColor(android.R.color.background_light))
                }
            }
            row.addView(tv)
        }
        tableLayout.addView(row)
    }
}
