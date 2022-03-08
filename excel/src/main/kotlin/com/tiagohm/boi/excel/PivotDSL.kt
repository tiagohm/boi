package com.tiagohm.boi.excel

import com.tiagohm.boi.core.BoiDSLMaker
import org.apache.poi.ss.SpreadsheetVersion
import org.apache.poi.ss.usermodel.DataConsolidateFunction
import org.apache.poi.ss.util.AreaReference
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFPivotTable
import org.apache.poi.xssf.usermodel.XSSFSheet

@BoiDSLMaker
class PivotDSL(val areaReference: String) {

    private val rows = mutableListOf<Int>()
    private val labels = mutableListOf<Triple<Int, String?, DataConsolidateFunction>>()

    var startingReference = "A1"

    fun row(index: Int) {
        rows.add(index)
    }

    fun column(
        index: Int,
        label: String? = null,
        fn: DataConsolidateFunction = DataConsolidateFunction.SUM,
    ) {
        labels.add(Triple(index, label, fn))
    }

    /**
     * I can't figure out to get this to work on lower levels than 0
     * https://stackoverflow.com/a/54182325/1832471
     */
    private fun XSSFPivotTable.collapseRows(level: Int) {
        val dataSheet = dataSheet as XSSFSheet
        val colAValues = LinkedHashSet<String>()

        for (r in 1 until dataSheet.lastRowNum + 1) {
            val row = dataSheet.getRow(r)

            if (row != null) {
                val cell = row.getCell(level)

                if (cell != null) {
                    colAValues.add(cell.toString())
                }
            }
        }

        val itemList = ctPivotTableDefinition.pivotFields.getPivotFieldArray(level).items.itemList
        var i = 0

        for (value in colAValues) {
            val item = itemList[i]
            item.unsetT()
            item.x = i++.toLong()

            pivotCacheDefinition.ctPivotCacheDefinition
                .cacheFields
                .getCacheFieldArray(level)
                .sharedItems.addNewS().v = value

            item.sd = false // set False will collapse child item
        }

        while (i < itemList.size) {
            val item = itemList[i++]
            item.sd = false // set False will collapse row.
        }
    }

    internal fun buildAndApply(sheet: XSSFSheet) {
        val aref = AreaReference(areaReference, SpreadsheetVersion.EXCEL2007)
        val pos = CellReference(startingReference)
        val pivotTable = sheet.createPivotTable(aref, pos)

        rows.map { pivotTable.addRowLabel(it) }

        labels.map { (index, label, fn) ->
            if (label == null) pivotTable.addColumnLabel(fn, index)
            else pivotTable.addColumnLabel(fn, index, label)
        }
    }
}