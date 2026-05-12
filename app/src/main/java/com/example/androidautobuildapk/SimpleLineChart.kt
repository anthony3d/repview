package com.example.repview

import android.content.Context
import android.graphics.*
import android.util.AttributeSet
import android.view.View
import java.text.SimpleDateFormat
import java.util.*

class SimpleLineChart @JvmOverloads constructor(
    context: Context,
    attrs: AttributeSet? = null,
    defStyleAttr: Int = 0
) : View(context, attrs, defStyleAttr) {
    
    private var dataPoints = mutableListOf<DataPoint>()
    private var movingAveragePoints = mutableListOf<DataPoint>()
    private var isDailyReport = false
    
    private val paintLine = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.rgb(76, 175, 80)
        strokeWidth = 4f
        style = Paint.Style.STROKE
    }
    
    private val paintAverageLine = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.rgb(255, 87, 34)
        strokeWidth = 3f
        style = Paint.Style.STROKE
    }
    
    private val paintPoint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.rgb(33, 150, 243)
        style = Paint.Style.FILL
    }
    
    private val paintText = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.BLACK
        textSize = 28f
    }
    
    private val paintGrid = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.LTGRAY
        strokeWidth = 1f
        style = Paint.Style.STROKE
    }
    
    private val paintWeekend = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.argb(80, 255, 193, 7)
        style = Paint.Style.FILL
    }
    
    private val dateFormat = SimpleDateFormat("dd.MM", Locale.getDefault())
    
    data class DataPoint(
        val date: Date,
        val value: Int
    )
    
    fun setData(data: List<DataPoint>, isDaily: Boolean = false) {
        dataPoints = data.toMutableList()
        isDailyReport = isDaily
        
        if (isDailyReport && dataPoints.size >= 3) {
            calculateMovingAverage()
        } else {
            movingAveragePoints.clear()
        }
        
        invalidate()
    }
    
    private fun calculateMovingAverage() {
        movingAveragePoints.clear()
        val windowSize = 7
        
        for (i in dataPoints.indices) {
            var sum = 0
            var count = 0
            val start = maxOf(0, i - windowSize / 2)
            val end = minOf(dataPoints.size - 1, i + windowSize / 2)
            
            for (j in start..end) {
                sum += dataPoints[j].value
                count++
            }
            
            val average = if (count > 0) sum / count else 0
            movingAveragePoints.add(DataPoint(dataPoints[i].date, average))
        }
    }
    
    private fun isWeekend(date: Date): Boolean {
        val calendar = Calendar.getInstance()
        calendar.time = date
        val dayOfWeek = calendar.get(Calendar.DAY_OF_WEEK)
        return dayOfWeek == Calendar.SATURDAY || dayOfWeek == Calendar.SUNDAY
    }
    
    private fun addDays(date: Date, days: Int): Date {
        val calendar = Calendar.getInstance()
        calendar.time = date
        calendar.add(Calendar.DAY_OF_YEAR, days)
        return calendar.time
    }
    
    override fun onDraw(canvas: Canvas) {
        super.onDraw(canvas)
        
        if (dataPoints.isEmpty()) {
            paintText.color = Color.GRAY
            paintText.textSize = 40f
            val text = "Нет данных для графика"
            val textWidth = paintText.measureText(text)
            canvas.drawText(
                text,
                width / 2 - textWidth / 2,
                height / 2.toFloat(),
                paintText
            )
            return
        }
        
        // Настройки отступов
        val paddingLeft = 80f
        val paddingRight = 40f
        val paddingTop = 40f
        val paddingBottom = 60f
        
        val chartWidth = width - paddingLeft - paddingRight
        val chartHeight = height - paddingTop - paddingBottom
        
        // Находим min и max значения
        val maxValue = dataPoints.maxOfOrNull { it.value } ?: 1
        val minValue = dataPoints.minOfOrNull { it.value } ?: 0
        val valueRange = if (maxValue == minValue) 1f else (maxValue - minValue).toFloat()
        
        // Рисуем фоновую подсветку выходных (только для ежедневного отчета)
        if (isDailyReport && dataPoints.size > 1) {
            var i = 0
            while (i < dataPoints.size) {
                val currentPoint = dataPoints[i]
                if (isWeekend(currentPoint.date)) {
                    // Находим начало выходных (суббота)
                    val startX = paddingLeft + (i * chartWidth / (dataPoints.size - 1).coerceAtLeast(1))
                    
                    // Находим конец выходных (воскресенье)
                    var endIndex = i
                    while (endIndex < dataPoints.size && isWeekend(dataPoints[endIndex].date)) {
                        endIndex++
                    }
                    endIndex--
                    
                    val endX = paddingLeft + ((endIndex + 1) * chartWidth / (dataPoints.size - 1).coerceAtLeast(1))
                    
                    // Закрашиваем область от субботы до понедельника (не включая понедельник)
                    canvas.drawRect(startX, paddingTop, endX, paddingTop + chartHeight, paintWeekend)
                    
                    i = endIndex + 1
                } else {
                    i++
                }
            }
        }
        
        // Рисуем сетку
        drawGrid(canvas, paddingLeft, paddingTop, chartWidth, chartHeight, maxValue, minValue)
        
        // Рисуем оси
        drawAxes(canvas, paddingLeft, paddingTop, chartWidth, chartHeight)
        
        // Рисуем точки и линии для основных данных
        val points = mutableListOf<Pair<Float, Float>>()
        
        for ((index, point) in dataPoints.withIndex()) {
            val x = paddingLeft + (index * chartWidth / (dataPoints.size - 1).coerceAtLeast(1))
            val y = paddingTop + chartHeight - ((point.value - minValue) / valueRange * chartHeight)
            
            points.add(x to y)
            
            // Рисуем точку
            canvas.drawCircle(x, y, 6f, paintPoint)
        }
        
        // Рисуем линию основных данных
        if (points.size > 1) {
            for (i in 0 until points.size - 1) {
                canvas.drawLine(
                    points[i].first, points[i].second,
                    points[i + 1].first, points[i + 1].second,
                    paintLine
                )
            }
        }
        
        // Рисуем линию скользящего среднего
        if (movingAveragePoints.isNotEmpty() && movingAveragePoints.size > 1) {
            val avgPoints = mutableListOf<Pair<Float, Float>>()
            
            for ((index, point) in movingAveragePoints.withIndex()) {
                val x = paddingLeft + (index * chartWidth / (movingAveragePoints.size - 1).coerceAtLeast(1))
                val y = paddingTop + chartHeight - ((point.value - minValue) / valueRange * chartHeight)
                avgPoints.add(x to y)
            }
            
            for (i in 0 until avgPoints.size - 1) {
                canvas.drawLine(
                    avgPoints[i].first, avgPoints[i].second,
                    avgPoints[i + 1].first, avgPoints[i + 1].second,
                    paintAverageLine
                )
            }
            
            // Рисуем точки скользящего среднего
            for (point in avgPoints) {
                canvas.drawCircle(point.first, point.second, 4f, paintAverageLine)
            }
        }
        
        // Рисуем даты на горизонтальной оси
        paintText.textSize = 22f
        paintText.color = Color.BLACK
        val step = maxOf(1, dataPoints.size / 10)
        
        for ((index, point) in dataPoints.withIndex()) {
            if (index % step == 0 || index == dataPoints.size - 1) {
                val x = paddingLeft + (index * chartWidth / (dataPoints.size - 1).coerceAtLeast(1))
                val dateText = dateFormat.format(point.date)
                val dateWidth = paintText.measureText(dateText)
                canvas.drawText(dateText, x - dateWidth / 2, height - paddingBottom + 25f, paintText)
            }
        }
        
        // Добавляем легенду для скользящего среднего
        if (movingAveragePoints.isNotEmpty()) {
            paintText.textSize = 20f
            paintText.color = Color.rgb(255, 87, 34)
            canvas.drawText("--- Скользящее среднее (7 дней)", width - 220f, 30f, paintText)
            
            paintText.color = Color.rgb(76, 175, 80)
            canvas.drawText("--- Данные", width - 220f, 55f, paintText)
        }
    }
    
    private fun drawGrid(
        canvas: Canvas,
        paddingLeft: Float,
        paddingTop: Float,
        chartWidth: Float,
        chartHeight: Float,
        maxValue: Int,
        minValue: Int
    ) {
        // Горизонтальные линии сетки (4 линии)
        for (i in 0..4) {
            val y = paddingTop + (i * chartHeight / 4)
            canvas.drawLine(paddingLeft, y, paddingLeft + chartWidth, y, paintGrid)
            
            // Подписи значений на вертикальной оси
            val value = maxValue - (i * (maxValue - minValue) / 4)
            paintText.color = Color.GRAY
            paintText.textSize = 24f
            val valueText = value.toInt().toString()
            canvas.drawText(valueText, 10f, y + 8f, paintText)
        }
        
        // Вертикальные линии сетки
        if (dataPoints.size > 1) {
            val step = maxOf(1, dataPoints.size / 10)
            for (i in 0 until dataPoints.size step step) {
                val x = paddingLeft + (i * chartWidth / (dataPoints.size - 1))
                canvas.drawLine(x, paddingTop, x, paddingTop + chartHeight, paintGrid)
            }
        }
    }
    
    private fun drawAxes(
        canvas: Canvas,
        paddingLeft: Float,
        paddingTop: Float,
        chartWidth: Float,
        chartHeight: Float
    ) {
        paintLine.color = Color.BLACK
        paintLine.strokeWidth = 2f
        
        // Ось Y
        canvas.drawLine(paddingLeft, paddingTop, paddingLeft, paddingTop + chartHeight, paintLine)
        
        // Ось X
        canvas.drawLine(paddingLeft, paddingTop + chartHeight, paddingLeft + chartWidth, paddingTop + chartHeight, paintLine)
    }
}
