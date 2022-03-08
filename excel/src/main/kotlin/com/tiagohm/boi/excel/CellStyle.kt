package com.tiagohm.boi.excel

import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.VerticalAlignment

data class CellStyle(
    val fillColor: IndexedColors? = null,
    val horizontalAlignment: HorizontalAlignment? = null,
    val verticalAlignment: VerticalAlignment? = null,
    val borderSettings: BorderRegion? = null,
    val font: Font = Font(),
    val wrapText: Boolean = false
)