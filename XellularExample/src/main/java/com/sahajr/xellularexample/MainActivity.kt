package com.sahajr.xellularexample

import android.annotation.SuppressLint
import android.os.Bundle
import android.support.v7.app.AppCompatActivity
import com.sahajr.xellular.Xellular
import kotlinx.android.synthetic.main.activity_main.*
import org.jetbrains.anko.doAsync
import org.jetbrains.anko.uiThread

class MainActivity : AppCompatActivity() {

    private var isReadWriteAllowed = true

    @SuppressLint("NewApi")
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        needsDynamicPermissions {
            isReadWriteAllowed = false
            requestPermissions(arrayOf(android.Manifest.permission.WRITE_EXTERNAL_STORAGE), 666)
        }

        createAndSave.setOnClickListener {
            if(!isReadWriteAllowed) {
                toast("Write access denied!")
            } else {
                // Initialize a Xellular object
                val x = Xellular()

                // Apache POI is slow enough on XLSX formats that it affects the UI Thread significantly.
                // Run it async
                doAsync {
                    val wb = x.workbook("tst") {
                        sheet("First") {
                            +"1||2||8||4||5"
                            +"a||b||c||d||e"
                        }
                        sheet("Second") {
                            row {
                                for(i in 1..5) {
                                    cell(i)
                                }
                            }
                        }
                        sheet("Third") {
                            header {
                                for(i in 1..5) {
                                    cell("Cell $i")
                                }
                            }
                            +"A||B||C||D||E"
                        }
                        sheet {
                            +"A||B||C||D||E"
                        }
                    }
                    uiThread {
                        wb.flushToDisk("mnt/sdcard") { success ->
                            if (success) toast("Successfully created the workbook!")
                            else toast("Failed to create workbook!")
                        }
                    }
                }

            }
        }
    }

    override fun onRequestPermissionsResult(requestCode: Int, permissions: Array<out String>, grantResults: IntArray) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        isReadWriteAllowed = true
    }

}
