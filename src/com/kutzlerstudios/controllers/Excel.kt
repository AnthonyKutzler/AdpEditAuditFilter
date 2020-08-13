package com.kutzlerstudios.controllers

import com.kutzlerstudios.objects.Person
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.util.*

class Excel {
    var book: XSSFWorkbook? = null
    private val file = File("/home/gob/Downloads/auditdoc.xlsx")
    private var person: Person? = null
    //var map = mutableMapOf<String, Person>()

    @Throws(Exception::class)
    fun runExcel() {
        book = XSSFWorkbook(FileInputStream(file))
        val sheet = book!!.getSheetAt(0)
        val rows = sheet.rowIterator()
        var map = mutableMapOf<String, Person>()

        val people: MutableList<Person> = ArrayList()
        var person = Person("TST","John", "Doe")
        var pid = "H1Q001"
        var y = 0
        var editor = ""

        while (rows.hasNext()) {
            println(y++)
            val row = rows.next()
            if(row.getCell(11) != null && !row.getCell(11).stringCellValue.contains("Position") && row.getCell(11).stringCellValue.trim() != ""){
                if(row.getCell(11).stringCellValue.contains("H") && row.getCell(11).stringCellValue != pid){
                    map[pid] = person
                    pid = row.getCell(11).stringCellValue
                    person = if(map[pid] != null)
                        map[pid]!!
                    else
                        Person(row.getCell(11).stringCellValue.substring(0..2), row.getCell(6).stringCellValue, row.getCell(0).stringCellValue)
                }
            }else{
                if(row.getCell(13) != null && row.getCell(13).stringCellValue.trim() != "" && !row.getCell(13).stringCellValue.contains("In ID")) {
                    if (row.getCell(13).stringCellValue == "TCMGR") {
                        person.addEdit()
                    } else {
                        person.addNonE()
                    }
                }
                if(row.getCell(16) != null && row.getCell(16).stringCellValue.trim() != "" && !row.getCell(16).stringCellValue.contains("Out ID")) {
                    if (row.getCell(16).stringCellValue == "TCMGR") {
                        person.addEdit()
                    } else {
                        person.addNonE()
                    }
                }
            }
            map[pid] = person




            /*if ((row.getCell(5) != null && !row.getCell(5).stringCellValue.contains("Last")) && row.getCell(5).cellTypeEnum != CellType.BLANK && !(row.getCell(5).stringCellValue.equals(person.lastName) && row.getCell(8).stringCellValue.equals(person.firstName))) {
                people.add(person)
                person = Person(row.getCell(0).stringCellValue.substring(0,3),row.getCell(8).stringCellValue, row.getCell(5).stringCellValue)
            }
            if (row.getCell(6) != null && row.getCell(6).stringCellValue.contains("Edit")) {
                editor = row.getCell(10).stringCellValue.split(",")[0].trim()
                if((row.lastCellNum > 12 && row.getCell(12) != null) && row.getCell(12).stringCellValue.contains("Time") && row.getCell(18).stringCellValue.contains("/")) {
                    person.addEdit()
                    person.addDate(
                        LocalDate.parse(
                            row.getCell(18).stringCellValue.split(" ".toRegex()).toTypedArray()[0],
                            DateTimeFormatter.ofPattern("MM/dd/yyyy")
                        ).toString() +
                                " (" + row.getCell(12).stringCellValue.split(" ")[1].trim() + ") - "// + editor
                    )
                }
            }else if(row.getCell(6) != null && row.getCell(6).stringCellValue.contains("Created")){
                person.addNonE()
            }*/
        }
        saveBook(map)
    }

    @Throws(Exception::class)
    private fun saveBook(people: Map<String, Person>) {
        val sheet = book!!.createSheet("new")
        var y = 0
        var x = 0
        var e = .0
        var nE = .0
        var tE = .0
        var tN = .0
        var co = "H1A"
        var z = 0
        val map = people.toSortedMap(compareBy<String> { it.length }.thenBy { it })
        for ((k,v) in map) {
            var row: Row = sheet.createRow(z)
            if(v.co != co){
                row.createCell(0).setCellValue(co)
                row.createCell(1).setCellValue(e/(e+nE))
                row.createCell(2).setCellValue(e)
                row.createCell(3).setCellValue(nE)
                row = sheet.createRow(z + (++y))
                tE += e
                tN += nE
                e = .0
                nE = .0
                co = v.co
            }
            row.createCell(0).setCellValue(v.co)
            row.createCell(1).setCellValue(v.firstName)
            row.createCell(2).setCellValue(v.lastName)
            row.createCell(3).setCellValue(v.getPercent())
            row.createCell(4).setCellValue(v.edit)
            row.createCell(5).setCellValue(v.nonE)
            e += v.edit
            nE += v.nonE
            /*    var y = 3;
            for(value in person!!.getDates().split(",")){
                row.createCell(y++).setCellValue(value)
            }
*/          x = z
            z++
        }
        var row: Row = sheet.createRow(x + 1)
        row.createCell(0).setCellValue(co)
        row.createCell(1).setCellValue(e/(e+nE))
        row.createCell(2).setCellValue(e)
        row.createCell(3).setCellValue(nE)
        tE += e
        tN += nE
        row = sheet.createRow(x + 2)
        row.createCell(0).setCellValue("TOTAL")
        row.createCell(1).setCellValue(tE/(tE+tN))
        row.createCell(2).setCellValue(tE)
        row.createCell(3).setCellValue(tN)

        book!!.write(FileOutputStream(file))
    }
}