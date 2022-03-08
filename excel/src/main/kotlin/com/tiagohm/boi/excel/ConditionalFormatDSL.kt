package com.tiagohm.boi.excel

import org.apache.poi.ss.usermodel.IndexedColors

data class ConditionalFormatDSL(
    var operator: ConditionalOperator,
    var formula: String,
    var fillColor: IndexedColors? = null,
)