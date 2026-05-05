package com.example.repview

import android.content.Intent
import android.graphics.Color
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
    private lateinit var lineChart: SimpleLineChart
    private val dateFormat = SimpleDateFormat("dd-MM-yyyy", Locale.getDefault())
    private val currentDate = Date()
    
    enum class ReportType {
        WEEKLY,  // Еженедельный (даты начала и конца недели разные)
        DAILY    // Ежедневный (даты начала и конца недели одинаковые)
    }
    
    data class WeekData(
        val startDate: Date,
        val endDate: Date,
        var sumValue: Int = 0
    )
    
    data class DailyData(
        val date: Date,
        val value: Int
    )
    
    // Храним прочитанные строки
    data class RawRow(
        val startDate: Date?,
        val endDate: Date?,
        val number: Int
    )
    
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_report_viewer)
        
        tableLayout = findViewById(R.id.tableLayout)
        lineChart = findViewById(R.id.lineChart)
        
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
            val inputStream = contentResolver.openInputStream(uri)
            val workbook = WorkbookFactory.create(inputStream)
            val sheet = workbook.getSheetAt(0)
            
            val headerRow = sheet.getRow(0)
            val lastColumnIndex = headerRow.lastCellNum - 1
            val preLastColumnIndex = lastColumnIndex - 1
            
            val startRow = 1
            val totalRows = sheet.physicalNumberOfRows
            val endRow = minOf(totalRows - 1, startRow + 99)
            
            // Сначала читаем все строки в память
            val rawRows = mutableListOf<RawRow>()
            for (i in startRow..endRow) {
                val row = sheet.getRow(i) ?: continue
                
                val startDateCell = row.getCell(2)
                val endDateCell = row.getCell(3)
                val numberCell = row.getCell(preLastColumnIndex)
                
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
                    rawRows.add(RawRow(startDate, endDate, number))
                }
            }
            
            workbook.close()
            inputStream?.close()
            
            if (rawRows.isEmpty()) {
                Toast.makeText(this, "Нет данных для отображения", Toast.LENGTH_SHORT).show()
                finish()
                return
            }
            
            // Определяем тип отчета по первым строкам
            val reportType = determineReportType(rawRows)
            
            if (reportType == ReportType.WEEKLY) {
                processWeeklyReport(rawRows)
            } else {
                processDailyReport(rawRows)
            }
            
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Ошибка: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }
    
    private fun determineReportType(rawRows: List<RawRow>): ReportType {
        val checkRows = rawRows.take(minOf(5, rawRows.size))
        if (checkRows.isEmpty()) return ReportType.WEEKLY
        
        val sameDateCount = checkRows.count { it.startDate == it.endDate }
        return if (sameDateCount > checkRows.size / 2) ReportType.DAILY else ReportType.WEEKLY
    }
    
    private fun processWeeklyReport(rawRows: List<RawRow>) {
        val weekMap = mutableMapOf<String, WeekData>()
        
        for (row in rawRows) {
            val startDate = row.startDate ?: continue
            val endDate = row.endDate ?: continue
            val number = row.number
            
            val key = "${dateFormat.format(startDate)}|${dateFormat.format(endDate)}"
            
            if (weekMap.containsKey(key)) {
                weekMap[key]?.sumValue = weekMap[key]!!.sumValue + number
            } else {
                weekMap[key] = WeekData(startDate, endDate, number)
            }
        }
        
        displayWeeklyTable(weekMap.values.toList())
    }
    
    private fun processDailyReport(rawRows: List<RawRow>) {
        val dailyDataList = mutableListOf<DailyData>()
        val weekGroups = mutableMapOf<Date, MutableList<Int>>()
        
        for (row in rawRows) {
            val date = row.startDate ?: continue
            val number = row.number
            
            dailyDataList.add(DailyData(date, number))
            
            val weekStart = getWeekStart(date)
            weekGroups.getOrPut(weekStart) { mutableListOf() }.add(number)
        }
        
        displayDailyTable(dailyDataList, weekGroups)
    }
    
    private fun getWeekStart(date: Date): Date {
        val calendar = Calendar.getInstance()
        calendar.time = date
        calendar.firstDayOfWeek = Calendar.MONDAY
        calendar.set(Calendar.DAY_OF_WEEK, Calendar.MONDAY)
        calendar.set(Calendar.HOUR_OF_DAY, 0)
        calendar.set(Calendar.MINUTE, 0)
        calendar.set(Calendar.SECOND, 0)
        calendar.set(Calendar.MILLISECOND, 0)
        return calendar.time
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
    
    private fun isCurrentDateInRange(startDate: Date, endDate: Date): Boolean {
        return currentDate in startDate..endDate
    }
    
    private fun displayWeeklyTable(weekDataList: List<WeekData>) {
        addWeeklyTableHeader()
        
        val chartData = mutableListOf<SimpleLineChart.DataPoint>()
        
        for (week in weekDataList) {
            val startDateStr = dateFormat.format(week.startDate)
            val endDateStr = dateFormat.format(week.endDate)
            val sumValue = week.sumValue.toString()
            val orderDate = addWeeks(week.startDate, 4)
            val paymentDate = addWeeks(week.startDate, 5)
            
            val isInRange = isCurrentDateInRange(orderDate, paymentDate)
            
            addWeeklyDataRow(
                startDateStr,
                endDateStr,
                sumValue,
                dateFormat.format(orderDate),
                dateFormat.format(paymentDate),
                isInRange
            )
            
            chartData.add(SimpleLineChart.DataPoint(paymentDate, week.sumValue))
        }
        
        lineChart.setData(chartData)
    }
    
    private fun displayDailyTable(dailyDataList: List<DailyData>, weekGroups: Map<Date, List<Int>>) {
        addDailyTableHeader()
        
        val chartData = mutableListOf<SimpleLineChart.DataPoint>()
        var currentWeekStart: Date? = null
        var weekOrderDate: String = ""
        var weekPaymentDate: String = ""
        
        for ((index, daily) in dailyDataList.withIndex()) {
            val weekStart = getWeekStart(daily.date)
            val dateStr = dateFormat.format(daily.date)
            val valueStr = daily.value.toString()
            
            if (currentWeekStart != weekStart) {
                currentWeekStart = weekStart
                weekOrderDate = dateFormat.format(addWeeks(weekStart, 4))
                weekPaymentDate = dateFormat.format(addWeeks(weekStart, 5))
            }
            
            val isInRange = isCurrentDateInRange(
                parseDateFromString(weekOrderDate) ?: Date(),
                parseDateFromString(weekPaymentDate) ?: Date()
            )
            
            val orderDisplay = if (index == 0 || getWeekStart(dailyDataList[index - 1].date) != weekStart) {
                weekOrderDate
            } else {
                ""
            }
            
            val paymentDisplay = if (index == 0 || getWeekStart(dailyDataList[index - 1].date) != weekStart) {
                weekPaymentDate
            } else {
                ""
            }
            
            addDailyDataRow(dateStr, valueStr, orderDisplay, paymentDisplay, isInRange)
            
            chartData.add(SimpleLineChart.DataPoint(daily.date, daily.value))
        }
        
        lineChart.setData(chartData)
    }
    
    private fun parseDateFromString(dateStr: String): Date? {
        return try {
            dateFormat.parse(dateStr)
        } catch (e: Exception) {
            null
        }
    }
    
    private fun addWeeklyTableHeader() {
        val headerRow1 = TableRow(this)
        
        val weekHeader = TextView(this).apply {
            text = "Неделя"
            setPadding(16, 12, 16, 12)
            setTextColor(Color.WHITE)
            setBackgroundColor(Color.rgb(33, 150, 243))
            textSize = 14f
            gravity = android.view.Gravity.CENTER
        }
        
        val weekHeaderSpan = TableRow.LayoutParams().apply {
            span = 2
        }
        weekHeader.layoutParams = weekHeaderSpan
        headerRow1.addView(weekHeader)
        
        val headers = arrayOf("Сумма", "Заказ", "Получка")
        for (header in headers) {
            val tv = TextView(this).apply {
                text = header
                setPadding(16, 12, 16, 12)
                setTextColor(Color.WHITE)
                setBackgroundColor(Color.rgb(33, 150, 243))
                textSize = 14f
                gravity = android.view.Gravity.CENTER
            }
            headerRow1.addView(tv)
        }
        tableLayout.addView(headerRow1)
        
        val headerRow2 = TableRow(this)
        val subHeaders = arrayOf("Начало", "Конец", "", "", "")
        
        for (subHeader in subHeaders) {
            val tv = TextView(this).apply {
                text = subHeader
                setPadding(16, 8, 16, 8)
                setTextColor(Color.WHITE)
                setBackgroundColor(Color.rgb(100, 181, 246))
                textSize = 12f
                gravity = android.view.Gravity.CENTER
            }
            headerRow2.addView(tv)
        }
        tableLayout.addView(headerRow2)
    }
    
    private fun addDailyTableHeader() {
        val headerRow = TableRow(this)
        
        val headers = arrayOf("Дата", "Значение", "Заказ", "Получка")
        for (header in headers) {
            val tv = TextView(this).apply {
                text = header
                setPadding(16, 12, 16, 12)
                setTextColor(Color.WHITE)
                setBackgroundColor(Color.rgb(33, 150, 243))
                textSize = 14f
                gravity = android.view.Gravity.CENTER
            }
            headerRow.addView(tv)
        }
        tableLayout.addView(headerRow)
    }
    
    private fun addWeeklyDataRow(col1: String, col2: String, col3: String, col4: String, col5: String, highlightColumns45: Boolean) {
        val row = TableRow(this)
        val data = arrayOf(col1, col2, col3, col4, col5)
        
        for ((index, value) in data.withIndex()) {
            val tv = TextView(this).apply {
                text = value
                setPadding(16, 12, 16, 12)
                textSize = 12f
                gravity = android.view.Gravity.CENTER
                setTextIsSelectable(true)
                
                when {
                    index < 2 -> {
                        setBackgroundColor(Color.WHITE)
                        setTextColor(Color.BLACK)
                    }
                    index == 2 -> {
                        setBackgroundColor(Color.rgb(255, 193, 7))
                        setTextColor(Color.BLACK)
                        textSize = 14f
                    }
                    index >= 3 -> {
                        if (highlightColumns45) {
                            setBackgroundColor(Color.rgb(76, 175, 80))
                            setTextColor(Color.WHITE)
                            textSize = 13f
                        } else {
                            if (index % 2 == 0) {
                                setBackgroundColor(Color.WHITE)
                            } else {
                                setBackgroundColor(Color.rgb(245, 245, 245))
                            }
                            setTextColor(Color.BLACK)
                        }
                    }
                }
            }
            row.addView(tv)
        }
        tableLayout.addView(row)
    }
    
    private fun addDailyDataRow(col1: String, col2: String, col3: String, col4: String, highlightColumns34: Boolean) {
        val row = TableRow(this)
        val data = arrayOf(col1, col2, col3, col4)
        
        for ((index, value) in data.withIndex()) {
            val tv = TextView(this).apply {
                text = value
                setPadding(16, 12, 16, 12)
                textSize = 12f
                gravity = android.view.Gravity.CENTER
                setTextIsSelectable(true)
                
                when {
                    index == 0 -> {
                        setBackgroundColor(Color.WHITE)
                        setTextColor(Color.BLACK)
                    }
                    index == 1 -> {
                        setBackgroundColor(Color.rgb(255, 193, 7))
                        setTextColor(Color.BLACK)
                        textSize = 14f
                    }
                    index >= 2 -> {
                        if (value.isNotEmpty() && highlightColumns34) {
                            setBackgroundColor(Color.rgb(76, 175, 80))
                            setTextColor(Color.WHITE)
                            textSize = 13f
                        } else {
                            setBackgroundColor(Color.WHITE)
                            if (value.isEmpty()) {
                                setTextColor(Color.WHITE) // Делаем пустые ячейки незаметными
                            } else {
                                setTextColor(Color.LTGRAY)
                            }
                        }
                    }
                }
            }
            row.addView(tv)
        }
        tableLayout.addView(row)
    }
}
