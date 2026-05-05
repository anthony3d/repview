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
    private val paintLine = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.rgb(76, 175, 80)
        strokeWidth = 4f
        style = Paint.Style.STROKE
    }
    
    private val paintPoint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.rgb(33, 150, 243)
        style = Paint.Style.FILL
    }
    
    private val paintText = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.BLACK
        textSize = 24f
    }
    
    private val paintGrid = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.LTGRAY
        strokeWidth = 1f
        style = Paint.Style.STROKE
    }
    
    private val dateFormat = SimpleDateFormat("dd.MM", Locale.getDefault())
    
    data class DataPoint(
        val date: Date,
        val value: Int
    )
    
    fun setData(data: List<DataPoint>) {
        dataPoints = data.toMutableList()
        invalidate() // Перерисовываем
    }
    
    override fun onDraw(canvas: Canvas) {
        super.onDraw(canvas)
        
        if (dataPoints.isEmpty()) {
            // Рисуем сообщение об отсутствии данных
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
        val paddingTop = 60f
        val paddingBottom = 80f
        
        val chartWidth = width - paddingLeft - paddingRight
        val chartHeight = height - paddingTop - paddingBottom
        
        // Находим min и max значения
        val maxValue = dataPoints.maxOfOrNull { it.value } ?: 1
        val minValue = dataPoints.minOfOrNull { it.value } ?: 0
        val valueRange = if (maxValue == minValue) 1f else (maxValue - minValue).toFloat()
        
        // Рисуем сетку
        drawGrid(canvas, paddingLeft, paddingTop, chartWidth, chartHeight, maxValue, minValue)
        
        // Рисуем оси
        drawAxes(canvas, paddingLeft, paddingTop, chartWidth, chartHeight)
        
        // Рисуем точки и линии
        val points = mutableListOf<Pair<Float, Float>>()
        
        for ((index, point) in dataPoints.withIndex()) {
            val x = paddingLeft + (index * chartWidth / (dataPoints.size - 1).coerceAtLeast(1))
            val y = paddingTop + chartHeight - ((point.value - minValue) / valueRange * chartHeight)
            
            points.add(x to y)
            
            // Рисуем точку
            canvas.drawCircle(x, y, 8f, paintPoint)
            
            // Рисуем значение над точкой
            paintText.color = Color.BLACK
            paintText.textSize = 28f
            val valueText = point.value.toString()
            val textWidth = paintText.measureText(valueText)
            canvas.drawText(valueText, x - textWidth / 2, y - 15f, paintText)
            
            // Рисуем дату под точкой
            paintText.textSize = 22f
            val dateText = dateFormat.format(point.date)
            val dateWidth = paintText.measureText(dateText)
            canvas.drawText(dateText, x - dateWidth / 2, y + 35f, paintText)
        }
        
        // Рисуем линию
        if (points.size > 1) {
            for (i in 0 until points.size - 1) {
                canvas.drawLine(
                    points[i].first, points[i].second,
                    points[i + 1].first, points[i + 1].second,
                    paintLine
                )
            }
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
            
            // Подписи значений
            val value = maxValue - (i * (maxValue - minValue) / 4)
            paintText.color = Color.GRAY
            paintText.textSize = 24f
            val valueText = value.toInt().toString()
            canvas.drawText(valueText, 10f, y + 8f, paintText)
        }
        
        // Вертикальные линии сетки (по количеству точек)
        if (dataPoints.size > 1) {
            for (i in 0 until dataPoints.size) {
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
        
        // Подписи осей
        paintText.color = Color.BLACK
        paintText.textSize = 28f
        
        // Подпись оси Y
        canvas.save()
        canvas.rotate(-90f, 25f, height / 2f)
        canvas.drawText("Сумма", 25f, height / 2f + 10f, paintText)
        canvas.restore()
        
        // Подпись оси X
        canvas.drawText("Дата получения", width / 2f - 80f, height - 20f, paintText)
    }
}
