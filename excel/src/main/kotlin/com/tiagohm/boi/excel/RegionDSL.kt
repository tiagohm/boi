package com.tiagohm.boi.excel

import com.tiagohm.boi.core.BoiDSLMaker
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

@BoiDSLMaker
class RegionDSL(
    val rowSpan: Int,
    val colSpan: Int,
) : Cell {

    internal val rows = mutableListOf<RowDSL>()

    fun row(block: RowDSL.() -> Unit) {
        RowDSL(rows.rowIndex).apply(block).let(rows::add)
    }

    fun emptyRow(count: Int = 1) {
        repeat(count) { rows.add(RowDSL(rows.rowIndex)) }
    }

    internal fun buildAndApply(
        workbook: XSSFWorkbook,
        sheet: XSSFSheet,
        startRowIndex: Int,
        startColIndex: Int,
    ): List<CellRangeAddress> {
        val actualRowSpan = rows.sumOf(RowDSL::span)

        require(actualRowSpan <= rowSpan) {
            // TODO: Provide some more information about where in the document this error happened, as this is all DSL we need better error reporting!
            "Number of rows within region '$actualRowSpan' when '$rowSpan' rows are required!"
        }

        val ranges = mutableListOf<CellRangeAddress>()

        var currentRowIndex = startRowIndex

        rows.map {
            sheet.getRow(currentRowIndex).run {
                val nestedRegions = it.buildAndApply(workbook, sheet, this, startColIndex)
                ranges.addAll(nestedRegions)
            }

            currentRowIndex += it.span
        }

        return ranges
    }
}