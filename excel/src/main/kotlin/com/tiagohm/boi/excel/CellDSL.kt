package com.tiagohm.boi.excel

import com.tiagohm.boi.core.BoiDSLMaker
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.*
import java.time.LocalDateTime
import java.util.Calendar
import java.util.Date

@BoiDSLMaker
open class CellDSL(
    private val parent: RowDSL,
    val currentColumnIndex: Int,
) : Cell {

    private val conditionalFormatting = mutableListOf<ConditionalFormatDSL>()

    var value: Any? = null
    var span = 1
    var fillColor = parent.fillColor
    var wrapText = parent.wrapText ?: false

    var verticalAlignment: VerticalAlignment? = null
    var horizontalAlignment: HorizontalAlignment? = null

    var borderSettings: BorderRegion? = null

    var font = parent.font ?: Font()
    var hyperlink: HyperLink? = null

    init {
        if (parent.borderStyle != null)
            border(
                style = parent.borderStyle!!.style,
                sides = parent.borderStyle!!.side,
                color = parent.borderStyle!!.color
            )
    }

    fun border(
        style: BorderStyle? = BorderStyle.THIN,
        sides: BorderSide = BorderSide.ALL,
        color: IndexedColors = IndexedColors.BLACK,
    ) {
        if (style == null) borderSettings = null
        else if (borderSettings == null) borderSettings = BorderRegion()

        if (borderSettings == null) return

        // TOP
        if (sides in listOf(BorderSide.TOP, BorderSide.TOP_BOTTOM, BorderSide.ALL)) {
            borderSettings!!.borderTop = style
            borderSettings!!.borderTopColor = color
        }

        // RIGHT
        if (sides in listOf(BorderSide.RIGHT, BorderSide.LEFT_RIGHT, BorderSide.ALL)) {
            borderSettings!!.borderRight = style
            borderSettings!!.borderRightColor = color
        }

        // BOTTOM
        if (sides in listOf(BorderSide.BOTTOM, BorderSide.TOP_BOTTOM, BorderSide.ALL)) {
            borderSettings!!.borderBottom = style
            borderSettings!!.borderBottomColor = color
        }

        // LEFT
        if (sides in listOf(BorderSide.LEFT, BorderSide.LEFT_RIGHT, BorderSide.ALL)) {
            borderSettings!!.borderLeft = style
            borderSettings!!.borderLeftColor = color
        }
    }

    fun font(block: Font.() -> Unit) {
        font = Font().apply(block)
    }

    fun hyperlink(block: HyperLink.() -> Unit) {
        hyperlink = HyperLink().apply(block)
    }

    fun addConditionalFormatting(
        operator: ConditionalOperator = ConditionalOperator.EQUAL,
        formula: String,
        color: IndexedColors? = null,
    ) {
        conditionalFormatting.add(ConditionalFormatDSL(operator, formula, color))
    }

    internal open fun buildAndApply(
        workbook: XSSFWorkbook,
        sheet: XSSFSheet,
        cell: XSSFCell,
    ) {
        cell.cellStyle = CellStyle(
            fillColor = fillColor,
            horizontalAlignment = horizontalAlignment,
            verticalAlignment = verticalAlignment,
            borderSettings = borderSettings,
            font = font,
            wrapText = wrapText
        ).getCachedStyle(workbook)

        hyperlink?.takeIf { it.address != null }?.also {
            cell.hyperlink = workbook.creationHelper
                .createHyperlink(it.type).apply { address = it.address }
        }

        conditionalFormatting.forEach { cf ->
            sheet.sheetConditionalFormatting.addConditionalFormatting(
                arrayOf(
                    CellRangeAddress(
                        parent.currentRow - 1,
                        parent.currentRow - 1,
                        currentColumnIndex,
                        currentColumnIndex
                    )
                ),
                sheet.sheetConditionalFormatting.createConditionalFormattingRule(cf.operator.byte, cf.formula).apply {
                    cf.fillColor?.let { color ->
                        createPatternFormatting().apply {
                            fillBackgroundColor = color.index
                        }
                    }
                },
            )
        }

        when (value) {
            null, "" -> return
            is String -> cell.setCellValue(value as String)
            is Number -> {
                cell.cellType = CellType.NUMERIC
                cell.setCellValue((value as Number).toDouble())
            }
            is Boolean -> cell.setCellValue(value as Boolean)
            is Date -> cell.setCellValue(value as Date)
            is LocalDateTime -> cell.setCellValue(value as LocalDateTime)
            is Calendar -> cell.setCellValue(value as Calendar)
            is CellFormula -> cell.cellFormula = (value as CellFormula).formula
            is XSSFRichTextString -> cell.setCellValue(value as XSSFRichTextString)
            else -> throw IllegalStateException("Type of value '$value' is not supported")
        }
    }

    internal fun Font.getCachedFont(workbook: XSSFWorkbook) = fontSet.getOrPut(workbook to this) {
        workbook.createFont().apply {
            this@getCachedFont.fontName?.let { this@apply.fontName = it }
            this@apply.fontHeightInPoints = this@getCachedFont.heightInPoints
            this@apply.bold = this@getCachedFont.bold
            this@apply.italic = this@getCachedFont.italic
            this@apply.strikeout = this@getCachedFont.strikeout
            this@getCachedFont.color?.let { this@apply.color = it.getIndex() }
        }
    }

    private fun CellStyle.getCachedStyle(workbook: XSSFWorkbook) = styleSet.getOrPut(workbook to this) {
        workbook.createCellStyle().apply {
            fillColor?.let {
                this@apply.fillForegroundColor = it.getIndex()
                this@apply.fillPattern = FillPatternType.SOLID_FOREGROUND
            }

            horizontalAlignment?.let { this@apply.alignment = it }
            verticalAlignment?.let { this@apply.verticalAlignment = it }

            borderSettings?.let { bs ->
                bs.borderTop?.let {
                    this@apply.borderTop = it
                    this@apply.topBorderColor = bs.borderTopColor!!.index
                }
                bs.borderRight?.let {
                    this@apply.borderRight = it
                    this@apply.rightBorderColor = bs.borderRightColor!!.index
                }
                bs.borderBottom?.let {
                    this@apply.borderBottom = it
                    this@apply.bottomBorderColor = bs.borderBottomColor!!.index
                }
                bs.borderLeft?.let {
                    this@apply.borderLeft = it
                    this@apply.leftBorderColor = bs.borderLeftColor!!.index
                }
            }

            setFont(this@getCachedStyle.font.getCachedFont(workbook))

            this@getCachedStyle.font.numberFormat?.let { nf ->
                if (CellDSL.dataFormat == null) CellDSL.dataFormat = workbook.createDataFormat()
                this@apply.dataFormat = CellDSL.dataFormat!!.getFormat(nf)
            }

            this@apply.wrapText = wrapText
        }
    }

    internal companion object {
        private val fontSet = mutableMapOf<Pair<XSSFWorkbook, Font>, XSSFFont>()
        private val styleSet = mutableMapOf<Pair<XSSFWorkbook, CellStyle>, XSSFCellStyle>()
        private var dataFormat: XSSFDataFormat? = null
    }
}