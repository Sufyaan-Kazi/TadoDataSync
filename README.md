# 🌡️ Tado Multi-Zone Smart Heating Data Sync (Enhanced Analytics)

> **Acknowledgment:** This project is a multi-zone fork of the original work by [theotocopulitos/TadoDataSync](https://github.com/theotocopulitos/TadoDataSync).

This project allows you to automatically capture every temperature change, boiler request, and weather fluctuation directly into a Google Sheet.

![Measured Data Dashboard](images/dashboard_sample.png)

Note: This is a fork of the original project https://github.com/theotocopulitos/TadoDataSync, where I have added multi-zone support and translated error messages from Spanish to English. Many thanks theotocopulitos for the original work in creating the base script :)

---

## ⚠️ CRITICAL: API Limits & Execution Time
If you do **not** have a paid Tado Auto-Assist subscription, Tado imposes a strict limit of **100 API calls per day**.
- **The Strategy:** We use a "Slow-Sync" method (downloading 10–20 days per night) to avoid being banned by Tado.
- **Execution Timeout:** Google Apps Script has a maximum runtime. If you set your historical batch too high (e.g., 100 days), the script will crash before finishing. 
- **Recommendation:** Keep `DAYS_PER_HISTORY_BATCH` between **10 and 30 days**. This is the "sweet spot" for reliability and speed.

---

## 🛠️ Phase 1: Preparation

1. Create a new [Google Sheet](https://sheets.new).
2. Open the Script Editor: **Extensions** > **Apps Script**.
3. Create the files:
   - Rename the default `Code.gs` to `Main.gs` and paste the **Main Script** code.
   - Click the **+** (Plus icon) > **Script**. Name it `Auth.gs` and paste the **Handshake Script** code.
4. **Save:** Click the floppy disk icon in the toolbar.

---

## 🔐 Phase 2: The "Handshake" (Safe Authorization)

Since this is a custom script, Google and Tado will ask for explicit permission. This is a one-time setup.

### Step A: Granting Google Permissions
1. In the top toolbar, select the function `requestAuthorization` and click **Run**.
2. A window titled **"Authorization Required"** will appear. Click **Review Permissions**.
3. Select your Google account.
4. **The "Unverified" Warning:** Because you just created this script, Google will warn you. 
   - Click **Advanced** (bottom left).
   - Click **Go to [Your Project Name] (unsafe)**.
5. Click **Allow**.

### Step B: Connecting to Tado
1. Look at the **Execution Log** at the bottom of the screen.
2. Find the link: `1. OPEN THIS LINK: https://login.tado.com/...`
3. **Copy and open that link** in your browser.
4. Log in to Tado and click **Approve**.
5. Return to the Google Script log. Copy the `device_code` (e.g., `_Xy7...`).

### Step C: Saving the "Master Key" (Refresh Token)
1. In `Auth.gs`, find `DEVICE_CODE` at the top and paste your code from Step B inside the quotes.
2. Select the function `finalizeTokenRequest` in the toolbar and click **Run**.
3. The log will show a **✅ SUCCESS!** message and a long **Refresh Token**. Copy that token.
4. Click the **Gear icon** (Project Settings) on the left sidebar.
5. Scroll down to **Script properties** and click **Add script property**.
6. Enter `INITIAL_REFRESH_TOKEN` as the Property, paste your token as the Value, and click **Save script properties**.

---

## 🏡 Phase 3: Finding your Home ID

The script needs to know which specific house to monitor.
1. Return to the code editor (`<>` icon). In the toolbar, select `discoverHomeId` and click **Run**.
2. The log will show: `🏡 Home: [Your Home Name] | ID: 123456`. Copy the numeric ID.
3. Go back to the **Gear icon** (Project Settings) > **Script properties**.
4. Click **Add script property**, enter `HOME_ID` as the Property, paste the numeric ID as the Value, and click **Save**.
*(Note: Ensure your `CLIENT_ID` is also saved here in Script properties if you haven't hardcoded it).*

---

## 🧪 Phase 4: Manual Testing (Verify before Automating)

Before setting the automatic timers, verify that everything is configured correctly:

1. **Test Midday Sync:** In the toolbar, select `todaySync` and click **Run**. Check your spreadsheet; you should see data from earlier today appearing in the four new tabs.
2. **Test Daily Sync:** In the toolbar, select `yesterdaySync` and click **Run**. Check your spreadsheet; you should see yesterday's data appearing in the four new tabs.
3. **Test Historical Sync:** In the toolbar, select `historicalSync` and click **Run**. Wait for it to finish. You should see about 10–20 days of past data added to the bottom of your sheets.
4. **Check for Errors:** If any function fails, double-check your `HOME_ID` and `INITIAL_REFRESH_TOKEN` in `Main.gs`.

---

## 📅 Phase 5: Automation (The "Slow-Sync" Triggers)

Once the manual tests are successful, set up two "alarms" (Triggers) to make it automatic:

1. **Daily Sync (Yesterday's Data):**
   - Click the **Alarm Clock icon** (Triggers) in the left sidebar.
   - Click **+ Add Trigger**.
   - Function: `yesterdaySync` | Source: `Time-driven` | Type: `Day timer` | Time: `1am to 2am`.
   - (Note: It is crucial to use yesterdaySync here so it safely grabs the entire previous 24-hour block after midnight)

2. **Historical Sync (The Past):**
   - Click **+ Add Trigger**.
   - Function: `historicalSync` | Source: `Time-driven` | Type: `Day timer` | Time: `4am to 5am`.
   - This will fetch 10-20 days every night automatically until your history limit date is reached.

---

## 📈 Phase 6: Monitoring Script Health

To ensure your sync never stops without you knowing:
1. **Executions Dashboard:** Click the **Executed Tasks (List icon)** on the left sidebar. You can see every time the script ran and if it succeeded.
2. **Error Notifications:** If a trigger fails, Google sends an email to your Gmail. If you see many "Failed" status logs, your Tado token might need a manual refresh via Phase 2.

---

## 📁 Where is my data?
The script automatically generates 4 tabs:
- **Measured Data:** Temperature and humidity every 15 minutes.
- **Settings:** Every change in target temperature (Setpoints).
- **Call For Heat:** Boiler activity and intensity levels.
- **Weather:** Sunlight and outdoor temperature logs.

## 🔧 Maintenance & Troubleshooting

### Resetting the History Start Point
If you want the script to "forget" its progress and restart the historical recovery from your `HISTORY_LIMIT_DATE`:

1.  **Via Code:** Run the `resetHistoryProgress` function from the Apps Script editor.
2.  **Via Settings:** - Click the **Gear icon** (Project Settings).
    - Scroll to **Script Properties**.
    - Delete the entry for `LAST_PROCESSED_DATE`.

### Handling Duplicates & Sorting
The script is now highly optimized to keep your spreadsheet fast and clean:
- **Idle Filtering:** Automatically drops "NONE" heat demand periods, keeping the spreadsheet lightweight.
- **Block-Wipe Prevention:** Uses a custom `clearDataForDate` mechanism to cleanly wipe a specific day's data right before syncing it, ensuring duplicate blocks are never stacked.
- **Auto-Sorts:** Keeps your sheets in perfect chronological order (newest at the top) regardless of when historical or daily data is synced.

### 🧹 Universal Cleanup Tool
If you ever suspect your sheet has duplicate data (e.g., from manual testing or older script versions), you can run a full sweep:
1. Open the Apps Script editor.
2. Select the `cleanAllSheetsEffective` function from the toolbar and click **Run**.
3. The script will scan all 4 tabs, remove any duplicate rows based on exact timestamps, and perfectly re-sort your data.
