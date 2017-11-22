package com.sahajr.xellularexample

import android.os.Build
import android.support.v7.app.AppCompatActivity
import android.widget.Toast

/**
 * Created by sahajr on Wed, 22/11/17.
 */
inline fun needsDynamicPermissions(requestPermissions: () -> Unit) {
    if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
        requestPermissions()
    }
}

fun AppCompatActivity.toast(message: String) {
    Toast.makeText(this, message, Toast.LENGTH_LONG).show()
}
