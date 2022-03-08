package com.tiagohm.boi.excel

import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class RichTextDSL(
    parent: RowDSL,
    currentColIndex: Int,
) : CellDSL(parent, currentColIndex) {

    private val texts = mutableListOf<RichTextIndexedDSL>()

    fun add(str: String, font: Font? = null) {
        texts.add(RichTextIndexedDSL(str, font))
    }

    override fun buildAndApply(
        workbook: XSSFWorkbook,
        sheet: XSSFSheet,
        cell: XSSFCell,
    ) {
        val richText = workbook.creationHelper.createRichTextString(texts.joinToString(separator = "") { it.text })

        var pointer = 0

        texts.forEach { config ->
            val end = pointer + config.text.length
            config.font?.let { richText.applyFont(pointer, end, it.getCachedFont(workbook)) }
            pointer = end
        }

        value = richText

        super.buildAndApply(workbook, sheet, cell)
    }

}