/** * TADO DATA SYNC - UNIFIED SHEET VERSION **/
const CONFIG = {
  HOME_ID: PropertiesService.getScriptProperties().getProperty("HOME_ID"), // PUT YOUR HOME_ID FROM AUTH_SCRIPT
  CLIENT_ID: PropertiesService.getScriptProperties().getProperty("CLIENT_ID"),
  INITIAL_REFRESH_TOKEN: PropertiesService.getScriptProperties().getProperty("INITIAL_REFRESH_TOKEN"),

  DAYS_PER_HISTORY_BATCH: 7, // ANYTHING FROM 10-30
  TIMEZONE: "GMT+1",
  DATE_FORMAT: "yyyy-MM-dd",
  TIME_FORMAT: "HH:mm:ss"
};

/**
 * Fetch todays data, can be used to get partial e.g. data so far today or the whole day depending on what time of day you run it
 */
function todaySync() {
  const today = new Date(new Date().getTime());

  dailySync(today);
}

/**
 * Fetch all of yesterdays data
 */
function yesterdaySync() {
  const today = new Date(new Date().getTime());
  const yesterday = new Date(today.getTime() - 24 * 60 * 60 * 1000);

  dailySync(yesterday);
}

/**
 * DAILY AUTOMATION
 */
function dailySync(dateToFetch) {
  const startTime = Date.now(); // ⬅️ 1. Start the timer

  validateConfig();
  const accessToken = getAccessToken();
  const headers = { 'Authorization': 'Bearer ' + accessToken };
  const zoneList = getZones(`${CONFIG.HOME_ID}`, headers);

  if (!dateToFetch) dateToFetch = new Date(new Date().getTime());
  const dateStr = Utilities.formatDate(dateToFetch, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);

  // Initialize daily batch arrays
  let dailyMeasured = [], dailySettings = [], dailyHeat = [], dailyWeather = [];

  for (let i = 1; i < zoneList.length; i++) {
    let zoneId = zoneList[i].id;
    let zoneName = zoneList[i].name;

    try {
      const zoneData = processAndSave(dateStr, "TOP", accessToken, zoneId, zoneName);
      if (zoneData) {
        dailyMeasured = dailyMeasured.concat(zoneData.measured);
        dailySettings = dailySettings.concat(zoneData.settings);
        dailyHeat = dailyHeat.concat(zoneData.heat);
        dailyWeather = dailyWeather.concat(zoneData.weather);
      }
    } catch (e) {
      console.error("❌ Error processing " + dateStr + " for " + zoneName + ": " + e.message);
    }
  }

  // BATCH SAVE: Write to the sheets exactly once for the entire day
  console.log(`💾 Saving batch: Measured(${dailyMeasured.length}), Settings(${dailySettings.length}), Heat(${dailyHeat.length}), Weather(${dailyWeather.length}) rows.`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  saveToSheet(ss, "Measured Data", ["ZONE ID", "ZONE NAME", "DATE", "TIME", "TEMP (C)", "HUM %", "SETPOINT", "HEATING", "EXT TEMP", "SOLAR %", "WEATHER", "SOURCE"], dailyMeasured);
  saveToSheet(ss, "Settings", ["ZONE ID", "ZONE NAME", "FROM", "TO", "TYPE", "SETPOINT", "SOURCE"], dailySettings);
  saveToSheet(ss, "Call For Heat", ["ZONE ID", "ZONE NAME", "FROM", "TO", "DEMAND", "SOURCE"], dailyHeat);
  saveToSheet(ss, "Weather", ["FROM", "TO", "STATE", "OUTSIDE TEMP", "SOLAR %", "SOURCE"], dailyWeather);

  // ⬅️ 2. Stop the timer and log the result
  const elapsedSecs = ((Date.now() - startTime) / 1000).toFixed(2);
  console.log(`⏱️ Total Elapsed Time: ${elapsedSecs} seconds`);
}

/**
 * HISTORICAL RECOVERY
 */
function historicalSync() {
  const startTime = Date.now(); // ⬅️ 1. Start the timer

  // 0. Validate configuration before starting
  validateConfig();

  const props = PropertiesService.getScriptProperties();
  const HISTORY_LIMIT_DATE = Utilities.formatDate(new Date(new Date().setDate(new Date().getDate() - CONFIG.DAYS_PER_HISTORY_BATCH)), "GMT+1", "yyyy-MM-dd");
  let lastDateStr = props.getProperty("LAST_PROCESSED_DATE");
  const limitDate = new Date(HISTORY_LIMIT_DATE);

  // Re-added standard JS multiplication format for the date math
  const yesterday = new Date(new Date().getTime() - 24 * 60 * 60 * 1000);

  console.log("LAST_PROCESSED_DATE: " + lastDateStr + ", HISTORY_LIMIT_DATE:" + limitDate);

  let lastDate;
  if (!lastDateStr || lastDateStr == null) {
    lastDate = new Date(limitDate.toDateString());
  } else {
    lastDate = new Date(lastDateStr);
    lastDate.setDate(lastDate.getDate() + 1);
  }
  if (lastDate < limitDate) {
    lastDate = limitDate;
  }

  if (lastDate > yesterday) {
    console.log("✅Historical data processed.")
    return;
  }

  const accessToken = getAccessToken();
  const headers = { 'Authorization': 'Bearer ' + accessToken };
  const zoneList = getZones(`${CONFIG.HOME_ID}`, headers);
  console.log("Starting historical recovery");

  // 1. OUTER LOOP: Days
  for (let dayIndex = 0; dayIndex < CONFIG.DAYS_PER_HISTORY_BATCH; dayIndex++) {

    const dateStr = Utilities.formatDate(lastDate, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);

    // Initialize batch arrays for this specific historical day
    let dailyMeasured = [], dailySettings = [], dailyHeat = [], dailyWeather = [];

    // 2. INNER LOOP: Process every zone for this specific date (except Hot Water - Zone 0)
    for (let z = 1; z < zoneList.length; z++) {
      let zoneId = zoneList[z].id;
      let zoneName = zoneList[z].name;

      try {
        const zoneData = processAndSave(dateStr, "HIST", accessToken, zoneId, zoneName);
        if (zoneData) {
          dailyMeasured = dailyMeasured.concat(zoneData.measured);
          dailySettings = dailySettings.concat(zoneData.settings);
          dailyHeat = dailyHeat.concat(zoneData.heat);
          dailyWeather = dailyWeather.concat(zoneData.weather);
        }
      } catch (e) {
        console.error("❌ Error processing " + dateStr + " for " + zoneName + ": " + e.message);
      }
    }

    // BATCH SAVE for the historical day: write to the sheet exactly once per tab
    console.log(`💾 Saving batch: Measured(${dailyMeasured.length}), Settings(${dailySettings.length}), Heat(${dailyHeat.length}), Weather(${dailyWeather.length}) rows.`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    saveToSheet(ss, "Measured Data", ["ZONE ID", "ZONE NAME", "DATE", "TIME", "TEMP (C)", "HUM %", "SETPOINT", "HEATING", "EXT TEMP", "SOLAR %", "WEATHER", "SOURCE"], dailyMeasured);
    saveToSheet(ss, "Settings", ["ZONE ID", "ZONE NAME", "FROM", "TO", "TYPE", "SETPOINT", "SOURCE"], dailySettings);
    saveToSheet(ss, "Call For Heat", ["ZONE ID", "ZONE NAME", "FROM", "TO", "DEMAND", "SOURCE"], dailyHeat);
    saveToSheet(ss, "Weather", ["FROM", "TO", "STATE", "OUTSIDE TEMP", "SOLAR %", "SOURCE"], dailyWeather);

    console.log("All zones processed and saved for: " + lastDate.toDateString());
    props.setProperty("LAST_PROCESSED_DATE", Utilities.formatDate(lastDate, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT));

    // 3. Move forward one day AFTER all zones are finished
    lastDate.setDate(lastDate.getDate() + 1);

    if (lastDate > yesterday) {
      console.log("✅Historical data processed.")
      break;
    }
  }

  // ⬅️ 2. Stop the timer and log the result
  const elapsedSecs = ((Date.now() - startTime) / 1000).toFixed(2);
  console.log(`⏱️ Total Elapsed Time: ${elapsedSecs} seconds`);
}

/**
 * UPDATED CORE PROCESSING LOGIC (With Boundary Filtering)
 */
function processAndSave(dateStr, mode, token, zoneId, zoneName) {
  console.log("\tFetching data for " + dateStr + " -> [" + zoneId + " / " + zoneName + "]");
  const data = fetchTadoData(dateStr, token, zoneId);

  // 1. Initial safety check
  if (!data || !data.measuredData) {
    console.warn("⚠️ No data available for zone: " + zoneName + " for the date: " + dateStr);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const toMs = (iso) => iso ? new Date(iso).getTime() : 0;
  const currentProcessingDate = new Date(dateStr).toDateString();

  // Extract blocks with null protection
  const temps = data.measuredData?.insideTemperature?.dataPoints || [];
  const hums = data.measuredData?.humidity?.dataPoints || [];
  const settings = data.settings?.dataIntervals || [];
  const heat = data.callForHeat?.dataIntervals || [];
  const weather = data.weather?.condition?.dataIntervals || [];
  const solar = data.weather?.solarIntensity?.dataPoints || [];

  // 2. GENERATE MASTER SHEET (With protection for each field)
  const unifiedRows = temps
    .filter(p => p.timestamp && new Date(p.timestamp).toDateString() === currentProcessingDate)
    .map(p => {
      const pMs = toMs(p.timestamp);
      const time = Utilities.formatDate(new Date(pMs), CONFIG.TIMEZONE, CONFIG.TIME_FORMAT);

      const hM = hums.reduce((prev, curr) => Math.abs(toMs(curr.timestamp) - pMs) < Math.abs(toMs(prev.timestamp) - pMs) ? curr : prev, hums[0]);
      const sM = settings.find(s => pMs >= toMs(s.from) && pMs < toMs(s.to));
      const heatM = heat.find(h => pMs >= toMs(h.from) && pMs < toMs(h.to));
      const wM = weather.find(w => pMs >= toMs(w.from) && pMs < toMs(w.to));
      const solarM = solar.reduce((prev, curr) => Math.abs(toMs(curr.timestamp) - pMs) < Math.abs(toMs(prev.timestamp) - pMs) ? curr : prev, solar[0]);

      return [
        zoneId, zoneName, dateStr, time,
        p.value?.celsius ?? "N/A",
        hM?.value ? (hM.value * 100).toFixed(1) : "N/A",
        sM?.value?.temperature?.celsius ?? "OFF",
        heatM?.value ?? "NONE",
        wM?.value?.temperature?.celsius ?? "N/A",
        solarM?.value?.percentage ? solarM.value.percentage + "%" : "0%",
        wM?.value?.state ?? "N/A",
        mode
      ];
    })
    .filter(row => row[7] !== "NONE");

  // 3. INTERVALS (Settings, Heat, Weather) with date protection
  const filterInterval = (item) => item && item.from && new Date(item.from).toDateString() === currentProcessingDate;

  const settingsRows = settings.filter(filterInterval).map(s => [
    zoneId, zoneName, formatTadoDate(s.from), formatTadoDate(s.to), s.value?.type || "N/A", s.value?.temperature?.celsius || "OFF", mode
  ]);

  const callForHeatRows = heat.filter(filterInterval).map(h => [
    zoneId, zoneName, formatTadoDate(h.from), formatTadoDate(h.to), h.value ?? "N/A", mode
  ]);

  const weatherRows = weather.filter(filterInterval).map(w => {
    const wMs = toMs(w.from);
    const solarM = solar.reduce((prev, curr) => Math.abs(toMs(curr.timestamp) - wMs) < Math.abs(toMs(prev.timestamp) - wMs) ? curr : prev, solar[0]);
    return [
      formatTadoDate(w.from), formatTadoDate(w.to), w.value?.state || "N/A", w.value?.temperature?.celsius || "N/A",
      solarM?.value?.percentage ? solarM.value.percentage + "%" : "0%", mode
    ];
  });

  // LOG ROW COUNTS PER ZONE (Indented to match the "Fetching" log)
  console.log(`\t  ↳ Extracted: Measured(${unifiedRows.length}), Settings(${settingsRows.length}), Heat(${callForHeatRows.length}), Weather(${weatherRows.length})`);

  // RETURN BATCH INSTEAD OF SAVING
  return {
    measured: unifiedRows,
    settings: settingsRows,
    heat: callForHeatRows,
    weather: weatherRows
  };
}