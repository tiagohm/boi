package com.tiagohm.boi.excel

import org.apache.poi.ss.usermodel.ComparisonOperator

enum class ConditionalOperator(val byte: Byte) {
    NO_COMPARISON(ComparisonOperator.NO_COMPARISON),
    BETWEEN(ComparisonOperator.BETWEEN),
    NOT_BETWEEN(ComparisonOperator.NOT_BETWEEN),
    EQUAL(ComparisonOperator.EQUAL),
    NOT_EQUAL(ComparisonOperator.NOT_EQUAL),
    GREATER_THAN(ComparisonOperator.GT),
    LESS_THAN(ComparisonOperator.LT),
    GREATER_THAN_OR_EQUAL_TO(ComparisonOperator.GE),
    LESS_THAN_OR_EQUAL_TO(ComparisonOperator.LE),
}