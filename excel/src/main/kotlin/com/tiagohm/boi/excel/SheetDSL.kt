package com.tiagohm.boi.excel

import com.tiagohm.boi.core.BoiDSLMaker
import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

@BoiDSLMaker
class SheetDSL(val sheetName: String? = null) {

    private val rows = mutableListOf<RowDSL>()
    private val columnConfigs = mutableListOf<ColumnConfig>()
    private var pivot: PivotDSL? = null

    fun row(rowCount: Int = 1, cellCount: Int = 0, block: RowDSL.() -> Unit = {}) {
        repeat(rowCount) {
            RowDSL(rows.rowIndex).apply {
                emptyCell(cellCount)
                block()
            }.let(rows::add)
        }
    }

    fun pivot(areaReference: String, block: PivotDSL.() -> Unit) {
        pivot = PivotDSL(areaReference).apply(block)
    }

    fun columnWidth(columnIndexes: List<Int>, widthSize: Int) {
        columnIndexes.map { columnWidth(it, widthSize) }
    }

    fun columnWidth(columnIndex: Int, widthSize: Int) {
        columnConfigs.add(ColumnConfig(columnIndex, widthSize, false))
    }

    fun autoColumnWidth(vararg columnIndexes: Int) {
        columnIndexes.map { autoColumnWidth(it) }
    }

    fun autoColumnWidth(columnIndex: Int) {
        columnConfigs.add(ColumnConfig(columnIndex, auto = true))
    }

    internal fun buildAndApply(workbook: XSSFWorkbook): XSSFSheet {
        val sheet = if (sheetName == null) workbook.createSheet()
        else workbook.createSheet(WorkbookUtil.createSafeSheetName(sheetName).trim())

        pivot?.buildAndApply(sheet)

        var currentIndex = 0

        val ranges = rows.flatMap { row ->
            sheet.createRow(currentIndex).let { r ->
                if (row.span > 1) repeat(row.span - 1) { sheet.createRow(currentIndex + it + 1) }
                currentIndex += row.span
                row.buildAndApply(workbook, sheet, r)
            }
        }

        // Add all ranges
        ranges.map { sheet.addMergedRegion(it) }

        pivot?.buildAndApply(sheet)

        // Set custom column widths
        columnConfigs.map { if (it.width != null && !it.auto) sheet.setColumnWidth(it.index, it.width) }
        columnConfigs.map { if (it.auto) sheet.autoSizeColumn(it.index) }

        return sheet
    }
}