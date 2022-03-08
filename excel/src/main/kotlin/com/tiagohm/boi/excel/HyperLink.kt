package com.tiagohm.boi.excel

import org.apache.poi.common.usermodel.HyperlinkType

data class HyperLink(
    var type: HyperlinkType = HyperlinkType.URL,
    var address: String? = null,
)
