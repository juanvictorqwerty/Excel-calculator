# Add project specific ProGuard rules here.
# You can control the set of applied configuration files using the
# proguardFiles setting in build.gradle.

# Apache POI
-keep class org.apache.poi.** { *; }
-keep class org.apache.xmlbeans.** { *; }
-keep class com.fasterxml.woodstox.** { *; }
-keep class org.openxmlformats.** { *; }
-keep class org.apache.commons.** { *; }
-keep class com.zaxxer.** { *; }

# Keep Apache POI methods
-keepclassmembers class org.apache.poi.** {
    *;
}

# Don't warn about Apache POI
-dontwarn org.apache.poi.**
-dontwarn org.apache.xmlbeans.**
-dontwarn org.apache.commons.**
-dontwarn com.fasterxml.woodstox.**
-dontwarn com.microsoft.schemas.**
-dontwarn org.openxmlformats.**
-dontwarn org.etsi.uri.**
-dontwarn com.zaxxer.**

# Keep model classes
-keep class com.example.excelcalculator.Student { *; }
-keep class com.example.excelcalculator.Student$* { *; }

# Keep Parcelable
-keepclassmembers class * implements android.os.Parcelable {
    public static final android.os.Parcelable$Creator *;
}
