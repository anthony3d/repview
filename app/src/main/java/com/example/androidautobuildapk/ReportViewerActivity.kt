package com.example.repview

import android.content.Intent
import android.graphics.Color
import android.net.Uri
import android.os.Bundle
import android.widget.LinearLayout
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
    private lateinit var dailyInfoBlock: LinearLayout
    private lateinit var weeklyInfoBlock: LinearLayout
    
    // Daily views
    private lateinit var thisWeekValue: TextView
    private lateinit var lastWeekValue: TextView
    private lateinit var thisMonthValue: TextView
    private lateinit var lastMonthValue: TextView
    
    // Weekly views
    private lateinit var last4WeeksValue: TextView
    private lateinit var prev4WeeksValue: TextView
    
    private val dateFormat = SimpleDateFormat("dd-MM-yyyy", Locale.getDefault())
    private val currentDate = Date()
    
    // Лимит в днях (300 дней ~ 10 месяцев)
    private val MAX_DAYS = 300
    
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
        dailyInfoBlock = findViewById(R.id.dailyInfoBlock)
        weeklyInfoBlock = findViewById(R.id.weeklyInfoBlock)
        
        // Daily views
        thisWeekValue = findViewById(R.id.thisWeekValue)
        lastWeekValue = findViewById(R.id.lastWeekValue)
        thisMonthValue = findViewById(R.id.thisMonthValue)
        lastMonthValue = findViewById(R.id.lastMonthValue)
        
        // Weekly views
        last4WeeksValue = findViewById(R.id.last4WeeksValue)
        prev4WeeksValue = findViewById(R.id.prev4WeeksValue)
        
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
            val endRow = totalRows - 1 // Читаем все строки
            
            // Сначала соберем все данные
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
            
            // Определяем тип отчета
            val reportType = if (allRowsData.take(5).count { it.first == it.second } > 2) {
                ReportType.DAILY
            } else {
                ReportType.WEEKLY
            }
            
            // Применяем лимит по дням
            val limitedData = applyDayLimit(allRowsData, reportType)
            
            if (limitedData.isEmpty()) {
                Toast.makeText(this, "Нет данных после применения лимита", Toast.LENGTH_SHORT).show()
                finish()
                return
            }
            
            // Находим последнюю дату в ограниченных данных
            val lastDate = limitedData.mapNotNull { it.first }.maxOrNull()
            
            if (reportType == ReportType.WEEKLY) {
                processWeeklyReport(limitedData, lastDate)
            } else {
                processDailyReport(limitedData, lastDate)
            }
            
            workbook.close()
            inputStream?.close()
            
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Ошибка: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }
    
    private fun applyDayLimit(allData: List<Triple<Date?, Date?, Int>>, reportType: ReportType): List<Triple<Date?, Date?, Int>> {
        if (allData.isEmpty()) return emptyList()
        
        // Находим все уникальные даты и сортируем их по убыванию
        val uniqueDates = allData.mapNotNull { it.first }.distinct().sortedDescending()
        
        if (uniqueDates.isEmpty()) return emptyList()
        
        // Берем последние MAX_DAYS дней
        val lastDates = uniqueDates.take(MAX_DAYS).toSet()
        
        // Фильтруем данные, оставляя только строки с датами из последних MAX_DAYS дней
        return allData.filter { 
            val date = it.first
            date != null && lastDates.contains(date)
        }
    }
    
    private fun processWeeklyReport(allRowsData: List<Triple<Date?, Date?, Int>>, lastDate: Date?) {
        // Фильтруем по последней дате
        val filteredData = allRowsData.filter { 
            val startDate = it.first
            lastDate == null || (startDate != null && !startDate.after(lastDate))
        }
        
        // Группируем по неделям
        val weekMap = mutableMapOf<String, WeekData>()
        
        for (data in filteredData) {
            val startDate = data.first ?: continue
            val endDate = data.second ?: continue
            val number = data.third
            
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
        
        val weekList = weekMap.values.toList()
        val sortedWeeks = weekList.sortedBy { it.startDate }
        
        calculateAndDisplayWeeklyStats(sortedWeeks)
        displayWeeklyTable(sortedWeeks)
    }
    
    private fun processDailyReport(allRowsData: List<Triple<Date?, Date?, Int>>, lastDate: Date?) {
        // Фильтруем по дате
        val filteredData = allRowsData.filter { 
            val startDate = it.first
            lastDate == null || (startDate != null && !startDate.after(lastDate))
        }
        
        // Группируем и суммируем по датам
        val dailySumMap = mutableMapOf<Date, Int>()
        val weekGroups = mutableMapOf<Date, MutableList<Int>>()
        
        for (data in filteredData) {
            val startDate = data.first ?: continue
            val number = data.third
            
            dailySumMap[startDate] = dailySumMap.getOrDefault(startDate, 0) + number
            
            val weekStart = getWeekStart(startDate)
            weekGroups.getOrPut(weekStart) { mutableListOf() }.add(number)
        }
        
        // Преобразуем в список и сортируем
        val dailyDataList = dailySumMap.map { (date, sum) -> DailyData(date, sum) }
            .sortedBy { it.date }
        
        if (dailyDataList.isEmpty()) {
            Toast.makeText(this, "Нет данных для отображения", Toast.LENGTH_SHORT).show()
            finish()
            return
        }
        
        calculateAndDisplayDailyStats(dailyDataList)
        displayDailyTable(dailyDataList, weekGroups)
    }
    
    // ... остальной код без изменений (getWeekStart, parseDate, addWeeks, isCurrentDateInRange,
    // calculateAndDisplayDailyStats, calculateAndDisplayWeeklyStats, displayWeeklyTable,
    // displayDailyTable, parseDateFromString, addWeeklyTableHeader, addDailyTableHeader,
    // addWeeklyDataRow, addDailyDataRow)
}
