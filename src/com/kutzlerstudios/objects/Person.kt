package com.kutzlerstudios.objects

import java.lang.StringBuilder

class Person(var co: String, var firstName: String, var lastName: String) {

    var edit: Double = .0
    var nonE: Double = .0

    var builder: StringBuilder = StringBuilder("")

    fun getDates() : String{
        return builder.toString()
    }

    fun addDate(value : String){
        builder.append("$value,")
    }

    fun addEdit(){
        edit++
    }

    fun addNonE(){
        nonE++
    }

    fun getPercent(): Double{
        if(edit > 0)
            return (edit/(edit + nonE))
        return .0
    }

    override fun toString(): String {
        return "$firstName $lastName"
    }
}