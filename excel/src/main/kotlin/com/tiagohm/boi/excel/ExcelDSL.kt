package com.tiagohm.boi.excel

import com.tiagohm.boi.core.BoiDSLMaker
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator
import org.apache.poi.xssf.usermodel.XSSFWorkbook

@BoiDSLMaker
class ExcelDSL {

    private val sheets = mutableListOf<SheetDSL>()
    private val sheetsOrder = mutableListOf<String>()

    var author = "Apache POI"
    var activeSheetIndex = 0 // TODO: Add support for this!

    fun sheet(sheetName: String? = null, block: SheetDSL.() -> Unit) {
        SheetDSL(sheetName).apply(block).let(sheets::add)
    }

    fun sheetsOrder(vararg sheetNames: String) {
        sheetsOrder.clear()
        sheetsOrder.addAll(sheetNames)
    }

    internal fun build(): XSSFWorkbook {
        val wb = XSSFWorkbook()

        wb.properties.apply {
            coreProperties.apply {
                creator = author
            }
        }

        sheets.forEach {
            it.buildAndApply(wb)
        }

        sheetsOrder.mapIndexed { index, sheetName ->
            wb.setSheetOrder(sheetName, index)
        }

        // Make sure all formulas are calculated
        XSSFFormulaEvaluator.evaluateAllFormulaCells(wb)

        return wb
    }
}