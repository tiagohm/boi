import com.tiagohm.boi.excel.excel
import io.kotest.core.spec.style.StringSpec
import java.time.LocalDateTime

class SimpleTest : StringSpec({
    beforeSpec {
        System.setProperty("java.awt.headless", "true")
    }

    "generate" {
        excel("./test.xlsx") {
            sheet("Planilha 1") {
                row {
                    cell(1.0)
                    cell("Text")
                    cellFormula("SUM(5.0 * 2.0)")
                    cell(56000.56) {
                        font { numberFormat = "[\$R\$ -416]#,##0.00" }
                    }
                    cell("Google") {
                        hyperlink { address = "https://google.com" }
                    }
                    cell(true)
                    cell(LocalDateTime.now()) {
                        font { numberFormat = "DD/MM/YY" }
                    }
                }

                (0..6).map { autoColumnWidth(it) }
            }
        }
    }
})