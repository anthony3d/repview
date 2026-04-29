package com.example.repview

import android.os.Bundle
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity

class MainActivity : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(TextView(this).apply {
            text = "Нажмите 'Поделиться' в вашем приложении с отчетами\nи выберите RepView"
            textSize = 18f
            setPadding(50, 50, 50, 50)
        })
    }
}
