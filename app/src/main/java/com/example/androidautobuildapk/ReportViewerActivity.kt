package com.example.repview

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.widget.TableLayout
import android.widget.TableRow
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.text.SimpleDateFormat
import java.util.*

class ReportViewerActivity : AppCompatActivity() {
    
    private lateinit var tableLayout: TableLayout
    private val dateFormat = SimpleDateFormat("dd-MM-yyyy", Locale.getDefault())
    
    // Класс для хранения данных по неделе
    data class WeekData(
        val startDate: Date,
        val endDate: Date,
        var sumValue: Int = 0
    )
    
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_report_viewer)
        
        tableLayout = findViewById(R.id.tableLayout)
        
        when {
            intent?.action == Intent.ACTION_SEND -> {
                val uri = intent.getParcelableExtra<Uri>(Intent.EXTRA_STREAM)
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
    
    private fun processFile(uri: Uri) {
        try {
            // Читаем файл напрямую из URI без создания временного файла
            val inputStream = contentResolver.openInputStream(uri)
            val workbook = WorkbookFactory.create(inputStream)
            val sheet = workbook.getSheetAt(0)
            
            // Получаем последний столбец
            val headerRow = sheet.getRow(0)
            val lastColumnIndex = headerRow.lastCellNum - 1
            val preLastColumnIndex = lastColumnIndex - 1
            
            // Собираем данные из строк
            val weekMap = mutableMapOf<String, WeekData>()
            
            // Определяем начальную строку (пропускаем заголовок)
            val startRow = 1
            // Берем последние 100 строк
            val totalRows = sheet.physicalNumberOfRows
            val endRow = minOf(totalRows - 1, startRow + 99)
            
            for (i in startRow..endRow) {
                val row = sheet.getRow(i) ?: continue
                
                // Получаем столбцы (индексация с 0)
                val startDateCell = row.getCell(2) // Столбец 3
                val endDateCell = row.getCell(3)   // Столбец 4
                val numberCell = row.getCell(preLastColumnIndex) // Предпоследний столбец
                
                val startDate = parseDate(startDateCell)
                val endDate = parseDate(endDateCell)
                val number = if (numberCell != null) {
                    try {
                        numberCell.numericCellValue.toInt()
                    } catch (e: Exception) {
                        0
                    }
                } else {
                    0
                }
                
                if (startDate != null && endDate != null) {
                    // Создаем ключ из дат
                    val key = "${dateFormat.format(startDate)}|${dateFormat.format(endDate)}"
                    
                    if (weekMap.containsKey(key)) {
                        // Суммируем
                        weekMap[key]?.sumValue = weekMap[key]!!.sumValue + number
                    } else {
                        // Добавляем новую запись
                        weekMap[key] = WeekData(startDate, endDate, number)
                    }
                }
            }
            
            workbook.close()
            inputStream?.close()
            
            if (weekMap.isEmpty()) {
                Toast.makeText(this, "Нет данных для отображения", Toast.LENGTH_SHORT).show()
                finish()
                return
            }
            
            // Отображаем таблицу
            displayTable(weekMap.values.toList())
            
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
                    val dateStr = cell.stringCellValue.trim()
                    val formats = listOf(
                        SimpleDateFormat("dd.MM.yyyy", Locale.getDefault()),
                        SimpleDateFormat("dd/MM/yyyy", Locale.getDefault()),
                        SimpleDateFormat("yyyy-MM-dd", Locale.getDefault()),
                        SimpleDateFormat("dd-MM-yyyy", Locale.getDefault())
                    )
                    for (format in formats) {
                        try {
                            return format.parse(dateStr)
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
    
    private fun displayTable(weekDataList: List<WeekData>) {
        addTableHeader()
        
        for (week in weekDataList) {
            val startDateStr = dateFormat.format(week.startDate)
            val endDateStr = dateFormat.format(week.endDate)
            val sumValue = week.sumValue.toString()
            val plus4Weeks = addWeeks(week.startDate, 4)
            val plus5Weeks = addWeeks(week.startDate, 5)
            
            addDataRow(
                startDateStr,
                endDateStr,
                sumValue,
                dateFormat.format(plus4Weeks),
                dateFormat.format(plus5Weeks)
            )
        }
    }
    
    private fun addTableHeader() {
        // Первая строка с общим заголовком для первых двух столбцов
        val headerRow1 = TableRow(this)
        
        // Создаем объединенный заголовок "Неделя" для первых двух столбцов
        val weekHeader = TextView(this).apply {
            text = "Неделя"
            setPadding(16, 12, 16, 12)
            setTextColor(resources.getColor(android.R.color.white))
            setBackgroundColor(resources.getColor(android.R.color.holo_blue_dark))
            textSize = 14f
            gravity = android.view.Gravity.CENTER
        }
        
        // Растягиваем на 2 столбца
        val weekHeaderSpan = TableRow.LayoutParams().apply {
            span = 2
        }
        weekHeader.layoutParams = weekHeaderSpan
        headerRow1.addView(weekHeader)
        
        // Остальные заголовки
        val headers = arrayOf("Сумма", "Заказ", "Получка")
        for (header in headers) {
            val tv = TextView(this).apply {
                text = header
                setPadding(16, 12, 16, 12)
                setTextColor(resources.getColor(android.R.color.white))
                setBackgroundColor(resources.getColor(android.R.color.holo_blue_dark))
                textSize = 14f
                gravity = android.view.Gravity.CENTER
            }
            headerRow1.addView(tv)
        }
        tableLayout.addView(headerRow1)
        
        // Вторая строка заголовков для подзаголовков
        val headerRow2 = TableRow(this)
        val subHeaders = arrayOf("Начало", "Конец", "", "", "")
        
        for (subHeader in subHeaders) {
            val tv = TextView(this).apply {
                text = subHeader
                setPadding(16, 8, 16, 8)
                setTextColor(resources.getColor(android.R.color.white))
                setBackgroundColor(resources.getColor(android.R.color.holo_blue_light))
                textSize = 12f
                gravity = android.view.Gravity.CENTER
            }
            headerRow2.addView(tv)
        }
        tableLayout.addView(headerRow2)
    }
    
    private fun addDataRow(col1: String, col2: String, col3: String, col4: String, col5: String) {
        val row = TableRow(this)
        val data = arrayOf(col1, col2, col3, col4, col5)
        
        for ((index, value) in data.withIndex()) {
            val tv = TextView(this).apply {
                text = value
                setPadding(16, 12, 16, 12)
                textSize = 12f
                gravity = android.view.Gravity.CENTER
                
                when {
                    index < 2 -> { // Первые два столбца (Неделя)
                        setBackgroundColor(resources.getColor(android.R.color.white))
                    }
                    index == 2 -> { // Сумма
                        setBackgroundColor(resources.getColor(android.R.color.holo_orange_light))
                        setTextColor(resources.getColor(android.R.color.black))
                        textSize = 14f
                    }
                    else -> { // Заказ и Получка
                        if (index % 2 == 0) {
                            setBackgroundColor(resources.getColor(android.R.color.white))
                        } else {
                            setBackgroundColor(resources.getColor(android.R.color.background_light))
                        }
                    }
                }
                setTextIsSelectable(true)
            }
            row.addView(tv)
        }
        tableLayout.addView(row)
    }
}
