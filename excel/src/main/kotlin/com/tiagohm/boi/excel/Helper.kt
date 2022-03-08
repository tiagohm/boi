package com.tiagohm.boi.excel

val Collection<RowDSL>.rowIndex
    get() = sumOf(RowDSL::span) + 1

val Collection<Cell>.cellIndex: Int
    get() = sumOf { cell ->
        when (cell) {
            is RegionDSL -> cell.rows.maxOf { it.cells.cellIndex }
            is CellDSL -> cell.span
        }
    }