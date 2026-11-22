package com.homeinventory.app;

import android.app.Activity;
import android.content.ActivityNotFoundException;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.os.IBinder;

/**
 * Opens the Home Inventory web app in a Chrome Custom Tab. Google blocks sign-in inside
 * plain WebViews, so the app runs in Chrome (sharing your Google login) with a green bar on top.
 */
public class MainActivity extends Activity {
    static final String APP_URL = "https://andrewmgooden.github.io/home-inventory/";
    static final int GREEN = 0xFF2E6A4C;

    private boolean launched;

    @Override
    protected void onCreate(Bundle state) {
        super.onCreate(state);
        launched = state != null && state.getBoolean("launched");
    }

    @Override
    protected void onResume() {
        super.onResume();
        // Back here after the tab was closed: close the app too.
        if (launched) {
            finish();
            return;
        }
        launched = true;
        Intent tab = new Intent(Intent.ACTION_VIEW, Uri.parse(APP_URL));
        Bundle extras = new Bundle();
        extras.putBinder("android.support.customtabs.extra.SESSION", (IBinder) null);
        tab.putExtras(extras);
        tab.putExtra("android.support.customtabs.extra.TOOLBAR_COLOR", GREEN);
        tab.putExtra("android.support.customtabs.extra.TITLE_VISIBILITY", 1);
        tab.setPackage("com.android.chrome");
        try {
            startActivity(tab);
        } catch (ActivityNotFoundException e) {
            tab.setPackage(null); // No Chrome: use whatever browser is installed.
            startActivity(tab);
        }
    }

    @Override
    protected void onSaveInstanceState(Bundle out) {
        super.onSaveInstanceState(out);
        out.putBoolean("launched", launched);
    }
}
