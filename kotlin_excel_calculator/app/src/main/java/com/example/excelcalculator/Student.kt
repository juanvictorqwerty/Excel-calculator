package com.example.excelcalculator

import android.os.Parcel
import android.os.Parcelable

data class Student(
    val name: String,
    val subjects: Map<String, Double>,
    val average: Double,
    val gpa: String
) : Parcelable {
    constructor(parcel: Parcel) : this(
        name = parcel.readString() ?: "",
        subjects = mutableMapOf<String, Double>().apply {
            val size = parcel.readInt()
            repeat(size) {
                val key = parcel.readString() ?: ""
                val value = parcel.readDouble()
                put(key, value)
            }
        },
        average = parcel.readDouble(),
        gpa = parcel.readString() ?: ""
    )

    override fun writeToParcel(parcel: Parcel, flags: Int) {
        parcel.writeString(name)
        parcel.writeInt(subjects.size)
        subjects.forEach { (key, value) ->
            parcel.writeString(key)
            parcel.writeDouble(value)
        }
        parcel.writeDouble(average)
        parcel.writeString(gpa)
    }

    override fun describeContents(): Int = 0

    companion object CREATOR : Parcelable.Creator<Student> {
        override fun createFromParcel(parcel: Parcel): Student = Student(parcel)
        override fun newArray(size: Int): Array<Student?> = arrayOfNulls(size)
    }
}
