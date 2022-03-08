package com.tiagohm.boi.excel

import com.tiagohm.boi.core.BoiDSLMaker
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.RegionUtil
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

@BoiDSLMaker
class RowDSL(val currentRow: Int) {

    data class CellBorder(
        val style: BorderStyle = BorderStyle.THIN,
        val side: BorderSide = BorderSide.ALL,
        val color: IndexedColors = IndexedColors.BLACK
    )

    internal val cells = mutableListOf<Cell>()

    var span = 1
    var heightInPoints = 15f

    // Default settings that will be used in columns if not overridden.
    var fillColor: IndexedColors? = null
    var borderStyle: CellBorder? = null
    var font: Font? = null
    var wrapText: Boolean? = null

    fun richCell(block: RichTextDSL.() -> Unit = {}) {
        RichTextDSL(this, cells.cellIndex)
            .apply(block)
            .let(cells::add)
    }

    fun cell(value: Any = "", block: CellDSL.() -> Unit = {}) {
        CellDSL(this, cells.cellIndex).apply {
            this.value = value
            block(this)
        }.let(cells::add)
    }

    fun cellFormula(formula: String, block: CellDSL.() -> Unit = {}) {
        CellDSL(this, cells.cellIndex).apply {
            this.value = CellFormula(formula)
            block(this)
        }.let(cells::add)
    }

    fun cellRegion(colspan: Int, block: RegionDSL.() -> Unit) {
        RegionDSL(rowSpan = span, colSpan = colspan).apply {
            block()
        }.let(cells::add)
    }

    fun emptyCell(count: Int = 1) {
        repeat(count) { cell("") }
    }

    internal fun buildAndApply(
        workbook: XSSFWorkbook,
        sheet: XSSFSheet,
        row: XSSFRow,
        startColIndex: Int = 0,
    ): List<CellRangeAddress> {
        val ranges = mutableListOf<CellRangeAddress>()
        var currentColIndex = startColIndex

        row.heightInPoints = heightInPoints

        cells.forEach { cell ->
            when (cell) {
                is CellDSL -> {
                    row.createCell(currentColIndex).let { cell.buildAndApply(workbook, sheet, it) }

                    if (span > 1 || cell.span > 1) {
                        val newRange = CellRangeAddress(
                            row.rowNum,
                            row.rowNum + span - 1,
                            currentColIndex,
                            currentColIndex + cell.span - 1
                        )

                        cell.borderSettings?.let { bs ->
                            bs.borderTop?.let { RegionUtil.setBorderTop(it, newRange, sheet) }
                            bs.borderTopColor?.let { RegionUtil.setTopBorderColor(it.index.toInt(), newRange, sheet) }
                            bs.borderRight?.let { RegionUtil.setBorderRight(it, newRange, sheet) }
                            bs.borderRightColor?.let {
                                RegionUtil.setRightBorderColor(
                                    it.index.toInt(),
                                    newRange,
                                    sheet
                                )
                            }
                            bs.borderBottom?.let { RegionUtil.setBorderBottom(it, newRange, sheet) }
                            bs.borderBottomColor?.let {
                                RegionUtil.setBottomBorderColor(
                                    it.index.toInt(),
                                    newRange,
                                    sheet
                                )
                            }
                            bs.borderLeft?.let { RegionUtil.setBorderLeft(it, newRange, sheet) }
                            bs.borderLeftColor?.let { RegionUtil.setLeftBorderColor(it.index.toInt(), newRange, sheet) }
                        }

                        ranges.add(newRange)
                    }

                    currentColIndex += cell.span
                }
                is RegionDSL -> {
                    val innerRegions = cell.buildAndApply(workbook, sheet, row.rowNum, currentColIndex)
                    ranges.addAll(innerRegions)
                    currentColIndex += cell.colSpan
                }
            }
        }

        return ranges
    }
}