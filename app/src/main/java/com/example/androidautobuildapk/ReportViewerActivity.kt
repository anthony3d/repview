package com.example.repview

import android.content.Intent
import android.graphics.Color
import android.net.Uri
import android.os.Bundle
import android.widget.LinearLayout
import android.widget.ScrollView
import android.widget.TableLayout
import android.widget.TableRow
import android.widget.TextView
import android.widget.Toast
import android.widget.ToggleButton
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.text.SimpleDateFormat
import java.util.*

class ReportViewerActivity : AppCompatActivity() {
    
    private lateinit var tableLayout: TableLayout
    private lateinit var lineChart: SimpleLineChart
    private lateinit var dailyInfoBlock: LinearLayout
    private lateinit var weeklyInfoBlock: LinearLayout
    private lateinit var tableScrollView: ScrollView
    private lateinit var amountModeToggle: ToggleButton
    
    // Daily views
    private lateinit var thisWeekValue: TextView
    private lateinit var lastWeekValue: TextView
    private lateinit var thisMonthValue: TextView
    private lateinit var lastMonthValue: TextView
    
    // Weekly views
    private lateinit var last4WeeksValue: TextView
    private lateinit var prev4WeeksValue: TextView
    
    private val dateFormat = SimpleDateFormat("dd-MM-yyyy", Locale.getDefault())
    
    private val currentDate: Date = Calendar.getInstance().apply {
        set(Calendar.HOUR_OF_DAY, 0)
        set(Calendar.MINUTE, 0)
        set(Calendar.SECOND, 0)
        set(Calendar.MILLISECOND, 0)
    }.time
    
    private val MAX_DAYS = 300
    
    // Индексы колонок в отчёте WB
    private val COL_START_DATE = 2
    private val COL_END_DATE = 3
    private val COL_SALE = 6          // G — «Продажа» (сколько заплатили покупатели)
    // Итого к оплате = предпоследняя колонка (T, индекс 19 при 21 колонке)
    
    enum class ReportType { WEEKLY, DAILY }
    enum class AmountMode { PAYMENT, SALE }
    
    // Одна строка отчёта: обе суммы сразу
    data class RowData(
        val startDate: Date,
        val endDate: Date,
        val paymentValue: Double,   // T — Итого к оплате
        val saleValue: Double       // G — Продажа
    ) {
        fun value(mode: AmountMode): Double =
            if (mode == AmountMode.SALE) saleValue else paymentValue
    }
    
    data class WeekData(
        val startDate: Date,
        val endDate: Date,
        var sumValue: Double = 0.0
    )
    
    data class DailyData(
        val date: Date,
        val value: Double
    )
    
    // Состояние: распарсенные данные и текущий режим
    private var allRowsData: List<RowData> = emptyList()
    private var reportType: ReportType = ReportType.DAILY
    private var lastDate: Date? = null
    private var amountMode: AmountMode = AmountMode.PAYMENT
    
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_report_viewer)
        
        tableLayout = findViewById(R.id.tableLayout)
        lineChart = findViewById(R.id.lineChart)
        dailyInfoBlock = findViewById(R.id.dailyInfoBlock)
        weeklyInfoBlock = findViewById(R.id.weeklyInfoBlock)
        tableScrollView = findViewById(R.id.tableScrollView)
        amountModeToggle = findViewById(R.id.amountModeToggle)
        
        thisWeekValue = findViewById(R.id.thisWeekValue)
        lastWeekValue = findViewById(R.id.lastWeekValue)
        thisMonthValue = findViewById(R.id.thisMonthValue)
        lastMonthValue = findViewById(R.id.lastMonthValue)
        
        last4WeeksValue = findViewById(R.id.last4WeeksValue)
        prev4WeeksValue = findViewById(R.id.prev4WeeksValue)
        
        // Тумблер: переключение режима
        amountModeToggle.setOnCheckedChangeListener { _, isChecked ->
            val newMode = if (isChecked) AmountMode.SALE else AmountMode.PAYMENT
            if (newMode != amountMode) {
                amountMode = newMode
                rebuildUi()
            }
        }
        
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
    
    // ===== Парсинг =====
    
    private fun parseDate(cell: org.apache.poi.ss.usermodel.Cell?): Date? {
        if (cell == null) return null
        return try {
            when (cell.cellType) {
                org.apache.poi.ss.usermodel.CellType.NUMERIC -> {
                    if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                        normalizeDate(cell.dateCellValue)
                    } else null
                }
                org.apache.poi.ss.usermodel.CellType.STRING -> {
                    val dateStr = cell.stringCellValue.trim()
                    val formats = listOf(
                        SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ssXXX", Locale.getDefault()),
                        SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss", Locale.getDefault()),
                        SimpleDateFormat("dd.MM.yyyy", Locale.getDefault()),
                        SimpleDateFormat("dd/MM/yyyy", Locale.getDefault()),
                        SimpleDateFormat("yyyy-MM-dd", Locale.getDefault()),
                        SimpleDateFormat("dd-MM-yyyy", Locale.getDefault())
                    )
                    for (format in formats) {
                        try {
                            val date = format.parse(dateStr) ?: continue
                            return normalizeDate(date)
                        } catch (e: Exception) { /* next */ }
                    }
                    null
                }
                else -> null
            }
        } catch (e: Exception) { null }
    }
    
    private fun normalizeDate(date: Date): Date {
        val calendar = Calendar.getInstance()
        calendar.time = date
        calendar.set(Calendar.HOUR_OF_DAY, 0)
        calendar.set(Calendar.MINUTE, 0)
        calendar.set(Calendar.SECOND, 0)
        calendar.set(Calendar.MILLISECOND, 0)
        return calendar.time
    }
    
    private fun readNumber(cell: org.apache.poi.ss.usermodel.Cell?): Double {
        if (cell == null) return 0.0
        return try {
            cell.numericCellValue
        } catch (e: Exception) { 0.0 }
    }
    
    private fun processFile(uri: Uri) {
        try {
            val inputStream = contentResolver.openInputStream(uri)
            val workbook = WorkbookFactory.create(inputStream)
            val sheet = workbook.getSheetAt(0)
            
            val headerRow = sheet.getRow(0)
            val lastColumnIndex = headerRow.lastCellNum - 1
            val preLastColumnIndex = lastColumnIndex - 1 // T — Итого к оплате
            
            val startRow = 1
            val totalRows = sheet.physicalNumberOfRows
            val endRow = totalRows - 1
            
            val parsed = mutableListOf<RowData>()
            
            for (i in startRow..endRow) {
                val row = sheet.getRow(i) ?: continue
                
                val startDate = parseDate(row.getCell(COL_START_DATE))
                val endDate = parseDate(row.getCell(COL_END_DATE))
                val saleValue = readNumber(row.getCell(COL_SALE))
                val paymentValue = readNumber(row.getCell(preLastColumnIndex))
                
                if (startDate != null && endDate != null) {
                    parsed.add(RowData(startDate, endDate, paymentValue, saleValue))
                }
            }
            
            workbook.close()
            inputStream?.close()
            
            if (parsed.isEmpty()) {
                Toast.makeText(this, "Нет данных для отображения", Toast.LENGTH_SHORT).show()
                finish()
                return
            }
            
            // Тип отчёта
            reportType = if (parsed.take(5).count { it.startDate == it.endDate } > 2) {
                ReportType.DAILY
            } else {
                ReportType.WEEKLY
            }
            
            // Лимит по дням
            val limited = applyDayLimit(parsed)
            if (limited.isEmpty()) {
                Toast.makeText(this, "Нет данных после применения лимита", Toast.LENGTH_SHORT).show()
                finish()
                return
            }
            
            allRowsData = limited
            lastDate = limited.mapNotNull { it.startDate }.maxOrNull()
            
            rebuildUi()
            
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Ошибка: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }
    
    private fun applyDayLimit(allData: List<RowData>): List<RowData> {
        if (allData.isEmpty()) return emptyList()
        val uniqueDates = allData.mapNotNull { it.startDate }.distinct().sortedDescending()
        if (uniqueDates.isEmpty()) return emptyList()
        val lastDates = uniqueDates.take(MAX_DAYS).toSet()
        return allData.filter { lastDates.contains(it.startDate) }
    }
    
    // ===== Построение UI (при старте и при переключении режима) =====
    
    private fun rebuildUi() {
        if (allRowsData.isEmpty()) return
        
        // Очищаем таблицу
        tableLayout.removeAllViews()
        
        val filtered = allRowsData.filter { row ->
            lastDate == null || !row.startDate.after(lastDate)
        }
        
        if (reportType == ReportType.WEEKLY) {
            buildWeeklyReport(filtered)
        } else {
            buildDailyReport(filtered)
        }
    }
    
    // Форматирование значения в зависимости от режима
    private fun formatValue(value: Double): String {
        return if (amountMode == AmountMode.SALE) {
            String.format(Locale.getDefault(), "%.2f", value)
        } else {
            Math.round(value).toString()
        }
    }
    
    // ===== Еженедельный отчёт =====
    
    private fun buildWeeklyReport(rows: List<RowData>) {
        // Группировка по нормализованным неделям
        val weekMap = mutableMapOf<String, WeekData>()
        
        for (row in rows) {
            val normalizedStart = getWeekStart(row.startDate)
            val normalizedEnd = getWeekEnd(row.endDate)
            val key = "${dateFormat.format(normalizedStart)}|${dateFormat.format(normalizedEnd)}"
            val value = row.value(amountMode)
            
            if (weekMap.containsKey(key)) {
                weekMap[key]!!.sumValue += value
            } else {
                weekMap[key] = WeekData(normalizedStart, normalizedEnd, value)
            }
        }
        
        if (weekMap.isEmpty()) {
            Toast.makeText(this, "Нет данных для отображения", Toast.LENGTH_SHORT).show()
            finish()
            return
        }
        
        val sortedWeeks = weekMap.values.sortedBy { it.startDate }
        calculateAndDisplayWeeklyStats(sortedWeeks)
        displayWeeklyTable(sortedWeeks)
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
    
    private fun getWeekEnd(date: Date): Date {
        val calendar = Calendar.getInstance()
        calendar.time = date
        calendar.firstDayOfWeek = Calendar.MONDAY
        calendar.set(Calendar.DAY_OF_WEEK, Calendar.SUNDAY)
        calendar.set(Calendar.HOUR_OF_DAY, 23)
        calendar.set(Calendar.MINUTE, 59)
        calendar.set(Calendar.SECOND, 59)
        calendar.set(Calendar.MILLISECOND, 999)
        return calendar.time
    }
    
    private fun addWeeks(date: Date, weeks: Int): Date {
        val calendar = Calendar.getInstance()
        calendar.time = date
        calendar.add(Calendar.WEEK_OF_YEAR, weeks)
        return calendar.time
    }
    
    private fun addDays(date: Date, days: Int): Date {
        val calendar = Calendar.getInstance()
        calendar.time = date
        calendar.add(Calendar.DAY_OF_YEAR, days)
        return calendar.time
    }
    
    private fun isCurrentDateInRange(startDate: Date, endDate: Date): Boolean {
        return currentDate in startDate..endDate
    }
    
    private fun calculateAndDisplayWeeklyStats(sortedWeeks: List<WeekData>) {
        weeklyInfoBlock.visibility = LinearLayout.VISIBLE
        dailyInfoBlock.visibility = LinearLayout.GONE
        
        if (sortedWeeks.size < 4) {
            last4WeeksValue.text = "Недостаточно данных"
            prev4WeeksValue.text = "Недостаточно данных"
            return
        }
        
        var last4WeeksSum = 0.0
        for (i in sortedWeeks.size - 4 until sortedWeeks.size) {
            last4WeeksSum += sortedWeeks[i].sumValue
        }
        
        var prev4WeeksSum = 0.0
        if (sortedWeeks.size >= 8) {
            for (i in sortedWeeks.size - 8 until sortedWeeks.size - 4) {
                prev4WeeksSum += sortedWeeks[i].sumValue
            }
        }
        
        last4WeeksValue.text = formatValue(last4WeeksSum)
        prev4WeeksValue.text = if (prev4WeeksSum > 0) formatValue(prev4WeeksSum) else "Нет данных"
    }
    
    private fun displayWeeklyTable(weekDataList: List<WeekData>) {
        addWeeklyTableHeader()
        
        val chartData = mutableListOf<SimpleLineChart.DataPoint>()
        val sortedWeeks = weekDataList.sortedBy { it.startDate }
        
        var previousWeekEnd: Date? = null
        val validWeeks = mutableListOf<WeekData>()
        for (week in sortedWeeks) {
            if (previousWeekEnd == null || week.startDate.after(previousWeekEnd) || week.startDate == previousWeekEnd) {
                validWeeks.add(week)
                previousWeekEnd = week.endDate
            }
        }
        
        for (week in validWeeks) {
            val startDateStr = dateFormat.format(week.startDate)
            val endDateStr = dateFormat.format(week.endDate)
            val sumValue = formatValue(week.sumValue)
            
            val orderDate = addWeeks(week.startDate, 3)
            val paymentDate = addDays(orderDate, 10)
            val isInRange = isCurrentDateInRange(orderDate, paymentDate)
            
            addWeeklyDataRow(
                startDateStr, endDateStr, sumValue,
                dateFormat.format(orderDate), dateFormat.format(paymentDate),
                isInRange
            )
            
            chartData.add(SimpleLineChart.DataPoint(paymentDate, week.sumValue))
        }
        
        lineChart.setData(chartData, false, amountMode == AmountMode.SALE)
        scrollTableToBottom()
    }
    
    // ===== Ежедневный отчёт =====
    
    private fun buildDailyReport(rows: List<RowData>) {
        val dailySumMap = mutableMapOf<Date, Double>()
        val weekGroups = mutableMapOf<Date, MutableList<Double>>()
        
        for (row in rows) {
            val value = row.value(amountMode)
            dailySumMap[row.startDate] = dailySumMap.getOrDefault(row.startDate, 0.0) + value
            val weekStart = getWeekStart(row.startDate)
            weekGroups.getOrPut(weekStart) { mutableListOf() }.add(value)
        }
        
        val filled = fillMissingDates(dailySumMap)
        if (filled.isEmpty()) {
            Toast.makeText(this, "Нет данных для отображения", Toast.LENGTH_SHORT).show()
            finish()
            return
        }
        
        calculateAndDisplayDailyStats(filled)
        displayDailyTable(filled, weekGroups)
    }
    
    private fun fillMissingDates(dailySumMap: Map<Date, Double>): List<DailyData> {
        if (dailySumMap.isEmpty()) return emptyList()
        val minDate = dailySumMap.keys.minOrNull() ?: return emptyList()
        val maxDate = dailySumMap.keys.maxOrNull() ?: return emptyList()
        
        val result = mutableListOf<DailyData>()
        val calendar = Calendar.getInstance()
        calendar.time = minDate
        
        while (!calendar.time.after(maxDate)) {
            val currentDate = calendar.time
            val value = dailySumMap[currentDate] ?: 0.0
            result.add(DailyData(currentDate, value))
            calendar.add(Calendar.DAY_OF_YEAR, 1)
        }
        return result
    }
    
    private fun calculateAndDisplayDailyStats(dailyDataList: List<DailyData>) {
        dailyInfoBlock.visibility = LinearLayout.VISIBLE
        weeklyInfoBlock.visibility = LinearLayout.GONE
        
        val thisWeekStart = getWeekStart(currentDate)
        val lastWeekStart = Calendar.getInstance().apply {
            time = thisWeekStart
            add(Calendar.WEEK_OF_YEAR, -1)
        }.time
        val thisMonthStart = Calendar.getInstance().apply {
            time = currentDate
            set(Calendar.DAY_OF_MONTH, 1)
            set(Calendar.HOUR_OF_DAY, 0); set(Calendar.MINUTE, 0)
            set(Calendar.SECOND, 0); set(Calendar.MILLISECOND, 0)
        }.time
        val lastMonthStart = Calendar.getInstance().apply {
            time = thisMonthStart
            add(Calendar.MONTH, -1)
        }.time
        
        var thisWeekSum = 0.0
        var lastWeekSum = 0.0
        var thisMonthSum = 0.0
        var lastMonthSum = 0.0
        
        for (data in dailyDataList) {
            if (data.date >= thisWeekStart) thisWeekSum += data.value
            if (data.date >= lastWeekStart && data.date < thisWeekStart) lastWeekSum += data.value
            if (data.date >= thisMonthStart) thisMonthSum += data.value
            if (data.date >= lastMonthStart && data.date < thisMonthStart) lastMonthSum += data.value
        }
        
        thisWeekValue.text = formatValue(thisWeekSum)
        lastWeekValue.text = formatValue(lastWeekSum)
        thisMonthValue.text = formatValue(thisMonthSum)
        lastMonthValue.text = formatValue(lastMonthSum)
    }
    
    private fun displayDailyTable(dailyDataList: List<DailyData>, weekGroups: Map<Date, List<Double>>) {
        addDailyTableHeader()
        
        val chartData = mutableListOf<SimpleLineChart.DataPoint>()
        var currentWeekStart: Date? = null
        var weekOrderDate = ""
        var weekPaymentDate = ""
        
        for ((index, daily) in dailyDataList.withIndex()) {
            val weekStart = getWeekStart(daily.date)
            val dateStr = dateFormat.format(daily.date)
            val valueStr = formatValue(daily.value)
            
            if (currentWeekStart != weekStart) {
                currentWeekStart = weekStart
                val orderDate = addWeeks(weekStart, 3)
                weekOrderDate = dateFormat.format(orderDate)
                val paymentDate = addDays(orderDate, 10)
                weekPaymentDate = dateFormat.format(paymentDate)
            }
            
            val orderDateObj = parseDateFromString(weekOrderDate)
            val paymentDateObj = parseDateFromString(weekPaymentDate)
            val isInRange = orderDateObj != null && paymentDateObj != null &&
                            isCurrentDateInRange(orderDateObj, paymentDateObj)
            
            val orderDisplay = if (index == 0 || getWeekStart(dailyDataList[index - 1].date) != weekStart) {
                weekOrderDate
            } else ""
            
            val paymentDisplay = if (index == 0 || getWeekStart(dailyDataList[index - 1].date) != weekStart) {
                weekPaymentDate
            } else ""
            
            addDailyDataRow(dateStr, valueStr, orderDisplay, paymentDisplay, isInRange)
            chartData.add(SimpleLineChart.DataPoint(daily.date, daily.value))
        }
        
        lineChart.setData(chartData, true, amountMode == AmountMode.SALE)
        scrollTableToBottom()
    }
    
    private fun parseDateFromString(dateStr: String): Date? {
        return try {
            dateFormat.parse(dateStr)
        } catch (e: Exception) { null }
    }
    
    private fun scrollTableToBottom() {
        tableScrollView.post {
            tableScrollView.fullScroll(ScrollView.FOCUS_DOWN)
        }
    }
    
    // ===== Заголовки и строки таблицы (без изменений) =====
    
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
        weekHeader.layoutParams = TableRow.LayoutParams().apply { span = 2 }
        headerRow1.addView(weekHeader)
        
        for (header in arrayOf("Сумма", "Заказ", "Получка")) {
            headerRow1.addView(TextView(this).apply {
                text = header
                setPadding(16, 12, 16, 12)
                setTextColor(Color.WHITE)
                setBackgroundColor(Color.rgb(33, 150, 243))
                textSize = 14f
                gravity = android.view.Gravity.CENTER
            })
        }
        tableLayout.addView(headerRow1)
        
        val headerRow2 = TableRow(this)
        for (subHeader in arrayOf("Начало", "Конец", "", "", "")) {
            headerRow2.addView(TextView(this).apply {
                text = subHeader
                setPadding(16, 8, 16, 8)
                setTextColor(Color.WHITE)
                setBackgroundColor(Color.rgb(100, 181, 246))
                textSize = 12f
                gravity = android.view.Gravity.CENTER
            })
        }
        tableLayout.addView(headerRow2)
    }
    
    private fun addDailyTableHeader() {
        val headerRow = TableRow(this)
        for (header in arrayOf("Дата", "Значение", "Заказ", "Получка")) {
            headerRow.addView(TextView(this).apply {
                text = header
                setPadding(16, 12, 16, 12)
                setTextColor(Color.WHITE)
                setBackgroundColor(Color.rgb(33, 150, 243))
                textSize = 14f
                gravity = android.view.Gravity.CENTER
            })
        }
        tableLayout.addView(headerRow)
    }
    
    private fun addWeeklyDataRow(col1: String, col2: String, col3: String, col4: String, col5: String, highlightColumns45: Boolean) {
        val row = TableRow(this)
        val data = arrayOf(col1, col2, col3, col4, col5)
        for ((index, value) in data.withIndex()) {
            row.addView(TextView(this).apply {
                text = value
                setPadding(16, 12, 16, 12)
                textSize = 12f
                gravity = android.view.Gravity.CENTER
                setTextIsSelectable(true)
                when {
                    index < 2 -> { setBackgroundColor(Color.WHITE); setTextColor(Color.BLACK) }
                    index == 2 -> { setBackgroundColor(Color.rgb(255, 193, 7)); setTextColor(Color.BLACK); textSize = 14f }
                    index >= 3 -> {
                        if (highlightColumns45) {
                            setBackgroundColor(Color.rgb(76, 175, 80)); setTextColor(Color.WHITE); textSize = 13f
                        } else {
                            setBackgroundColor(if (index % 2 == 0) Color.WHITE else Color.rgb(245, 245, 245))
                            setTextColor(Color.BLACK)
                        }
                    }
                }
            })
        }
        tableLayout.addView(row)
    }
    
    private fun addDailyDataRow(col1: String, col2: String, col3: String, col4: String, highlightColumns34: Boolean) {
        val row = TableRow(this)
        val data = arrayOf(col1, col2, col3, col4)
        for ((index, value) in data.withIndex()) {
            row.addView(TextView(this).apply {
                text = value
                setPadding(16, 12, 16, 12)
                textSize = 12f
                gravity = android.view.Gravity.CENTER
                setTextIsSelectable(true)
                when {
                    index == 0 -> { setBackgroundColor(Color.WHITE); setTextColor(Color.BLACK) }
                    index == 1 -> { setBackgroundColor(Color.rgb(255, 193, 7)); setTextColor(Color.BLACK); textSize = 14f }
                    index >= 2 -> {
                        if (value.isNotEmpty() && highlightColumns34) {
                            setBackgroundColor(Color.rgb(76, 175, 80)); setTextColor(Color.WHITE); textSize = 13f
                        } else {
                            setBackgroundColor(Color.WHITE)
                            setTextColor(if (value.isEmpty()) Color.TRANSPARENT else Color.LTGRAY)
                        }
                    }
                }
            })
        }
        tableLayout.addView(row)
    }
}
