package com.tiagohm.boi.excel

data class ColumnConfig(
    val index: Int,
    val width: Int? = null,
    val auto: Boolean = false,
)