# 🏠 Gooden's Home Inventory

Made by Andrew Gooden, BSc.

Keep track of what's in the house, let your shopping list build itself, tick off your monthly bills,
and see where the money goes. Everything is saved in your own Google Sheet, with a colorful dashboard
and charts.

- **Open the app:** https://andrewmgooden.github.io/home-inventory/
- **Try it with sample data (nothing is saved):** https://andrewmgooden.github.io/home-inventory/?demo

## How it fits together

```
 index.html on GitHub Pages ──(Google sign-in token)──▶ Apps Script web app (Code.gs) ──▶ your Google Sheet
                                                                                      └──▶ photos in Google Drive
```

- **The app page** (`index.html`, `manifest.webmanifest`, `icons/`) is hosted here on GitHub Pages.
- **The server** (`Code.gs`, `appsscript.json`) runs in Google Apps Script, attached to the Sheet.
  It only answers requests carrying a valid Google sign-in for the Sheet's owner, or for an email
  listed under `household` in the Sheet's **Settings** tab (comma-separated).
- Google sign-in uses an OAuth client in the "Home Inventory" Google Cloud project. Only that client's
  **test users** can sign in, and only from `https://andrewmgooden.github.io`.

## Using it on Android

- **Best:** open the app link in **Chrome** → **⋮** → **Install app** (or **Add to Home screen**).
  It installs as a full-screen app with the green house icon.
- **Or** install the APK (`android/`, built with `android/build-apk.ps1`). It opens the app in Chrome
  with a green bar on top. Keep `release.keystore` and `keystore-password.txt` from
  `%USERPROFILE%\home-inventory-android`, because Android only installs updates signed with the same key.

## Making changes

- **App page:** edit `index.html`, then commit and push. GitHub Pages updates in about a minute.
- **Server:** edit `Code.gs`, then run:

  ```
  npx @google/clasp@3.4.1 push --force
  npx @google/clasp@3.4.1 update-deployment AKfycbzspCiuOW4naWn9FoSe12v2aIjV-41pkKa0DbEYyse-tn_FfAR7OzYTk4qceyEWiRbLGA
  ```

  The second command keeps the same server address, so the app page doesn't need changing.
  If a change needs new Google permissions, open the script editor, run `setup` once, and approve.

## How to use it

| Tab | What it does |
|---|---|
| **At home** | Everything you own, grouped by category. Tap **+ / −** as you use things up. Tap an item to edit it, set its price, or add a photo. |
| **Shopping** | Anything at or below its **Restock at** number shows up here by itself. You can add one-off things at the top. Tap ✓ when you buy something, enter the price, and it's added to your stock and logged as spending. **Share list** sends the list by text or to your notes. |
| **Bills** | Light, water, phone, rent, car gas, insurance, and anything else you pay regularly. Each bill shows when it's due, turns red when it's overdue, and has a **Pay** button. Things you pay a little at a time, like car gas, go under **Pay as you go**. Bills that come every few months or once a year only show in the months they're due. |
| **Spending** | Everything you spent this month, with bills and shopping shown separately, broken down by category and compared with past months. Set a monthly budget and the currency (J$ by default) with the ⚙️ button. |

**Photos:** tap the photo box on an item to take a picture or pick one. Photos are shrunk, saved in a
Drive folder called *Home Inventory Photos*, and also shown in the Sheet's **Photo** column.

### The Sheet

- **Dashboard** (first tab): 12 summary cards, 6 charts (spending by month, by category, bills paid,
  stores, what you keep at home), this month's bills with ✅ Paid / ⏳ Due / ⚠️ Overdue, items running low,
  latest purchases, and what you buy most. It's all formulas, so don't type there. If it ever looks
  wrong, use **🏠 Inventory → Rebuild dashboard & colors**.
- **Inventory, Shopping List, Bills, Purchases, Settings**: color-coded tabs you can also edit by hand.
  Rows that need restocking turn red; bill payments in Purchases are shaded purple. Don't rename the
  column headers, because the app finds columns by name.

## Questions

**Can my partner or family use it?** Add their Google email under `household` in the Sheet's
**Settings** tab, and as a **test user** in the Google Cloud project (Google Auth Platform → Audience).

**It says my sign-in expired.** Google sign-ins last about an hour. The app asks again, usually with a
single tap.

**A photo doesn't show.** Photos are shared as "anyone with the link can view" so the app can show them.
Some work or school Google accounts don't allow that. A personal Gmail account works fine.
