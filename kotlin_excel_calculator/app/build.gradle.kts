plugins {
    id("com.android.application")
    id("org.jetbrains.kotlin.android")
}

android {
    namespace = "com.example.excelcalculator"
    compileSdk = 34

    defaultConfig {
        applicationId = "com.example.excelcalculator"
        minSdk = 26
        targetSdk = 34
        versionCode = 1
        versionName = "1.0"
    }

    buildTypes {
        release {
            isMinifyEnabled = false
            proguardFiles(
                getDefaultProguardFile("proguard-android-optimize.txt"),
                "proguard-rules.pro"
            )
        }
        debug {
            isMinifyEnabled = false
            proguardFiles(
                getDefaultProguardFile("proguard-android-optimize.txt"),
                "proguard-rules.pro"
            )
        }
    }
    
    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }
    
    kotlinOptions {
        jvmTarget = "17"
    }
    
    buildFeatures {
        viewBinding = true
    }
}

dependencies {
    implementation("androidx.core:core-ktx:1.12.0")
    implementation("androidx.appcompat:appcompat:1.6.1")
    implementation("com.google.android.material:material:1.11.0")
    implementation("androidx.constraintlayout:constraintlayout:2.1.4")
    implementation("androidx.recyclerview:recyclerview:1.3.2")
    
    // Apache POI for Excel reading
    implementation("org.apache.poi:poi:5.2.5")
    implementation("org.apache.poi:poi-ooxml:5.2.5")
    implementation("org.apache.poi:poi-ooxml-lite:5.2.5")
    implementation("org.apache.xmlbeans:xmlbeans:5.1.1")
    implementation("com.fasterxml.woodstox:woodstox-core:6.5.1")
    
    // Required Apache POI dependencies
    implementation("org.apache.commons:commons-collections4:4.4")
    implementation("org.apache.commons:commons-math3:3.6.1")
    implementation("org.apache.commons:commons-compress:1.26.0")
    implementation("commons-codec:commons-codec:1.16.0")
    implementation("commons-io:commons-io:2.15.1")
    implementation("com.zaxxer:SparseBitSet:1.3")
    
    // For file picker
    implementation("androidx.documentfile:documentfile:1.0.1")
}
