package com.sahajr.xellularexample

import android.app.Application
import com.sahajr.xellular.Xellular

/**
 * Created by sahajr on Wed, 22/11/17.
 */
class MainApplication: Application() {
    override fun onCreate() {
        super.onCreate()

        // Initialize Xellular
        Xellular.setup()
    }
}
