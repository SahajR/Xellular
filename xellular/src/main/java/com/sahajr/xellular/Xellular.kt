/**
 * Xellular - a DSL for creating xlsx documents in android
 *
 * Copyright (C) 2017 Sahaj Ramachandran
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

package com.sahajr.xellular

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.jetbrains.anko.doAsync
import org.jetbrains.anko.uiThread
import java.io.File
import java.io.FileOutputStream

class Xellular {

    companion object {
        /**
         * This setup is a workaround for <a href="http://poi.apache.org/faq.html#faq-N101E6">this</a>
         */
        fun setup() {
            System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl")
            System.setProperty("org.apache.poi.javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl")
            System.setProperty("org.apache.poi.javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl")
        }
    }

    @DslMarker
    annotation class WorkbookMarker

    /**
     * Represents a styleable cell that can hold a value of any standard data type.
     * @param value The value the cell holds
     * @param style The cell style conforming to XSSFCellStyle
     */
    @WorkbookMarker
    class XCell<out T: Any>(private val value: T,
                            private var style: XSSFCellStyle? = null) {
        // TODO: Find a good way to let users customize cell styles and pass them here
        private var type: CellType = CellType.BLANK
        init {
            when(value){
                is String -> type = CellType.STRING
                is Int -> type = CellType.NUMERIC
                is Float -> type = CellType.NUMERIC
                is Boolean -> type = CellType.BOOLEAN
            }
        }
        fun setStyle(style: XSSFCellStyle) {
            this.style = style
        }
        fun getStyle(): XSSFCellStyle? = style
        fun getType(): CellType = type
        fun getValue(): T = value
    }

    /**
     * A Row represents a collection of styleable cells.
     */
    @WorkbookMarker
    abstract class Row {
        abstract val cells: MutableList<XCell<Any>>
        abstract fun <T: Any> cell(value: T)
    }

    class XRow: Row() {
        override val cells = mutableListOf<XCell<Any>>()
        override fun <T: Any> cell(value: T) {
            val newCell = XCell<Any>(value)
            cells.add(newCell)
        }
    }

    /**
     * Headers are Rows that have the default header style applied to their cells
     */
    class XHeader(private val headerStyle: XSSFCellStyle): Row() {
        override val cells = mutableListOf<XCell<Any>>()
        override fun <T: Any> cell(value: T) {
            val newCell = XCell<Any>(value)
            newCell.setStyle(headerStyle)
            cells.add(newCell)
        }
    }

    /**
     *  A sheet in the workbook. This needs to have a reference to the workbook because of the way POI is defined.
     *  @param wb The reference to the workbook the sheet is created
     *  @param {String=} name An optional name given to the sheet
     */
    @WorkbookMarker
    class XSheet(private val wb: XWorkbook, private val name: String? = null): XSSFSheet() {
        private val rows = mutableListOf<Row>()

        private fun <T: Row> initRow(r: T, init: T.() -> Unit) {
            r.init()
            rows.add(r)
        }

        /**
         * For simple cells that hold pure string values, you can just use `+` followed by a string with the cell values delimited by `||`
         * <pre>
         *     {@code
         *          ...
         *          {
         *              row{
         *                  +"Cell1||Cell2||Cell3||Cell4||Cell5"
         *              }
         *          }
         *          ...
         *     }
         * </pre>
         */
        operator fun String.unaryPlus() {
            val cellValues = this.split("||")
            val newRow = XRow()
            cellValues.forEach { value ->
                newRow.cell(value)
            }
            rows.add(newRow)
        }

        fun header(init: XHeader.() -> Unit) = initRow(XHeader(wb.getStyle("DEFAULT_HEADER_STYLE")), init)
        fun row(init: XRow.() -> Unit) = initRow(XRow(), init)

        fun getName(): String? = name
        fun getRows(): List<Row> = rows
    }

    /**
     * An XSSFWorkbook that also holds default styles
     */
    @WorkbookMarker
    class XWorkbook(private val fileName: String): XSSFWorkbook() {

        private val styles = mutableMapOf<String, XSSFCellStyle>()
        fun getStyle(id: String): XSSFCellStyle = styles[id]!!

        private var sheetCount = 1

        // Add default header style
        init {
            styles["DEFAULT"] = this.createCellStyle()
            val headerStyle = this.createCellStyle()
            headerStyle.fillForegroundColor = IndexedColors.AQUA.index
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
            styles["DEFAULT_HEADER_STYLE"] = headerStyle
        }

        /**
         * Generates a sheet based on the rows and cells specified. The name is optional.
         */
        fun sheet(name: String? = null, init: XSheet.() -> Unit) {
            val xSheet = XSheet(this, name)
            xSheet.init()

            val newSheet = this.createSheet(xSheet.getName() ?: "Sheet ${sheetCount}")
            sheetCount++
            for((rowCount, row) in xSheet.getRows().withIndex()) {

                val newRow = newSheet.createRow(rowCount)
                for((cellCount, cell) in row.cells.withIndex()) {

                    val newCell = newRow.createCell(cellCount, cell.getType())
                    newCell.cellStyle = cell.getStyle() ?: getStyle("DEFAULT")
                    when(cell.getType()) {
                        CellType.NUMERIC -> {
                            when(cell.getValue()) {
                                is Float -> newCell.setCellValue((cell.getValue()as Float).toDouble())
                                is Int -> newCell.setCellValue((cell.getValue() as Int).toDouble())
                            }
                        }
                        CellType.STRING -> newCell.setCellValue(cell.getValue() as String)
                        CellType.BLANK -> newCell.setCellValue(cell.getValue() as String)
                        CellType.BOOLEAN -> newCell.setCellValue(cell.getValue() as Boolean)
                        else -> throw Exception("Illegal type encountered.")
                    }

                }

            }

        }

        /**
         * Writes the workbook to the disk at the specified location
         * @param filePath The location to dump the workbook file
         */
        fun flushToDisk(filePath: String, onSuccess: (success: Boolean) -> Unit) {
            doAsync {
                try {
                    val directory = File(filePath)
                    val file = File(directory, "$fileName.xlsx")
                    this@XWorkbook.write(FileOutputStream(file))
                    uiThread {
                        onSuccess(true)
                    }
                } catch (e: Exception) {
                    uiThread {
                        onSuccess(false)
                    }
                }
            }
        }

    }

    /**
     * Generates a workbook with the specified file name
     */
    fun workbook(fileName: String, init: XWorkbook.() -> Unit): XWorkbook {
        val newWorkbook = XWorkbook(fileName)
        newWorkbook.init()
        return newWorkbook
    }
}
