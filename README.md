![Xellular banner](art/xellular-banner.jpg "Xellular banner")
# Xellular
A DSL for creating XLSX documents in android.

## What it is for
As a wrapper around Apache POI, you can write xlsx documents in a declarative syntax.
```kotlin
val x = Xellular()

val wb = x.workbook("Example") {

    // Simple sheet
    sheet {
        +"A||B||C||D||E"
    }

    // Pure string row cells
    sheet("Second") {
        +"1||2||8||4||5"
        +"a||b||c||d||e"
    }

    // Rows can have cells that are of any standard data type
    sheet("Third") {
        row {
            for(i in 1..5) {
                cell(i)
            }
        }
    }

    // Rows can be styled as headers
    sheet("Fourth") {
        header {
            for(i in 1..5) {
                cell("Cell $i")
            }
        }
        +"A||B||C||D||E"
    }
}

// Once completed, you can write directly to disk
wb.flushToDisk("mnt/sdcard") { success ->
  if (success) toast("Successfully created the workbook!")
  else toast("Failed to create workbook!")
}
```

## Including it in your project
Include the jitpack repository to your project level gradle config.
```gradle
allprojects {
    repositories {
        ...
        maven { url 'https://jitpack.io' }
    }
}
```
And then in your app gradle config add:
```gradle
dependencies {
    implementation 'com.github.SahajR:Xellular:0.2.0'
}
```
Once you sync your project, you need to add this line in your code as early as possible. Preferably in `onCreate` of your Application.
```kotlin
class MainApplication: Application() {
    override fun onCreate() {
        super.onCreate()

        // Initialize Xellular
        Xellular.setup()
    }
}
```
This is required as a workaround for [this](http://poi.apache.org/faq.html#faq-N101E6) problem with POI.

## Ceveats
POI is generally slow in writing files of xlsx format. This may cause frame drops if you create workbooks on the main thread. Do them in async.
```kotlin
// Using anko doAsync
doAsync {
    val wb = x.workbook {...}
    uiThread {
        wb.flushToDisk("path"){...}
    }
}
```
See the [example](/XellularExample/src/main/java/com/sahajr/xellularexample/MainActivity.kt#L30).

## Todos
- [ ] Add support for custom cell styles
- [ ] Add support to write an arbitrary cell out of sequence
- [ ] Add support for Streaming workbooks (SXSSF) as they are faster than XSSF Workbooks
- [ ] Improve general performance

## License
```
The MIT License (MIT)

Copyright © 2017 SahajR

Permission is hereby granted, free of charge, to any person
obtaining a copy of this software and associated documentation
files (the “Software”), to deal in the Software without
restriction, including without limitation the rights to use,
copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the
Software is furnished to do so, subject to the following
conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.
```
