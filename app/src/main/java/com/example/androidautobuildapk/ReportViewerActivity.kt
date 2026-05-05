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
        WEEKLY,
        DAILY
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
            
            // Сначала соберем все данные для определения типа отчета и последней даты
            val allRowsData = mutableListOf<Triple<Date?, Date?, Int>>()
            
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
                    allRowsData.add(Triple(startDate, endDate, number))
                }
            }
            
            if (allRowsData.isEmpty()) {
                Toast.makeText(this, "Нет данных для отображения", Toast.LENGTH_SHORT).show()
                finish()
                return
            }
            
            // Находим последнюю дату
            val lastDate = allRowsData.maxOfOrNull { it.first }
            
            // Определяем тип отчета
            val reportType = if (allRowsData.take(5).count { it.first == it.second } > 2) {
                ReportType.DAILY
            } else {
                ReportType.WEEKLY
            }
            
            if (reportType == ReportType.WEEKLY) {
                processWeeklyReport(allRowsData, lastDate)
            } else {
                processDailyReport(allRowsData, lastDate)
            }
            
            workbook.close()
            inputStream?.close()
            
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Ошибка: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }
    
    private fun processWeeklyReport(allRowsData: List<Triple<Date?, Date?, Int>>, lastDate: Date?) {
        val weekMap = mutableMapOf<String, WeekData>()
        
        for (data in allRowsData) {
            val startDate = data.first ?: continue
            val endDate = data.second ?: continue
            val number = data.third
            
            // Пропускаем строки с датами после последней даты
            if (lastDate != null && startDate.after(lastDate)) {
                continue
            }
            
            val key = "${dateFormat.format(startDate)}|${dateFormat.format(endDate)}"
            
            if (weekMap.containsKey(key)) {
                weekMap[key]?.sumValue = weekMap[key]!!.sumValue + number
            } else {
                weekMap[key] = WeekData(startDate, endDate, number)
            }
        }
        
        if (weekMap.isEmpty()) {
            Toast.makeText(this, "Нет данных для отображения", Toast.LENGTH_SHORT).show()
            finish()
            return
        }
        
        displayWeeklyTable(weekMap.values.toList())
    }
    
    private fun processDailyReport(allRowsData: List<Triple<Date?, Date?, Int>>, lastDate: Date?) {
        val dailyDataList = mutableListOf<DailyData>()
        val weekGroups = mutableMapOf<Date, MutableList<Int>>()
        
        for (data in allRowsData) {
            val startDate = data.first ?: continue
            val number = data.third
            
            // Пропускаем строки с датами после последней даты
            if (lastDate != null && startDate.after(lastDate)) {
                continue
            }
            
            dailyDataList.add(DailyData(startDate, number))
            
            val weekStart = getWeekStart(startDate)
            weekGroups.getOrPut(weekStart) { mutableListOf() }.add(number)
        }
        
        // Сортируем по дате
        dailyDataList.sortBy { it.date }
        
        if (dailyDataList.isEmpty()) {
            Toast.makeText(this, "Нет данных для отображения", Toast.LENGTH_SHORT).show()
            finish()
            return
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
                            val date = format.parse(dateStr)
                            val calendar = Calendar.getInstance()
                            calendar.time = date
                            calendar.set(Calendar.HOUR_OF_DAY, 0)
                            calendar.set(Calendar.MINUTE, 0)
                            calendar.set(Calendar.SECOND, 0)
                            calendar.set(Calendar.MILLISECOND, 0)
                            return calendar.time
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
        val sortedWeeks = weekDataList.sortedBy { it.startDate }
        
        for (week in sortedWeeks) {
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
            
            val orderDateObj = parseDateFromString(weekOrderDate)
            val paymentDateObj = parseDateFromString(weekPaymentDate)
            val isInRange = orderDateObj != null && paymentDateObj != null && 
                            isCurrentDateInRange(orderDateObj, paymentDateObj)
            
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
                                setTextColor(Color.TRANSPARENT)
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
