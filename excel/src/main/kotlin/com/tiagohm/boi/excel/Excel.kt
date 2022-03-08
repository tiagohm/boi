package com.tiagohm.boi.excel

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File

fun excel(block: ExcelDSL.() -> Unit): XSSFWorkbook {
    return ExcelDSL().apply(block).build()
}

fun excel(
    path: String,
    block: ExcelDSL.() -> Unit
): File {
    val workbook = excel(block)

    return File(path).also {
        it.outputStream().use(workbook::write)
    }
}