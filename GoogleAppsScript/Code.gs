/************************************************************
 * EV SAFE GUARD V2
 *
 * Google Apps Script
 *
 * FEATURES
 * ----------------------------------------------------------
 * ESP32 / ESP8266 API (HTTP GET & POST)
 * Real-time Dashboard
 * Safety Threshold
 * Risk Score 0-100
 * 2 Minute Prediction
 * Anomaly Detection
 * Alarm History
 * LINE Notification
 * Thai AI Analysis
 * Historical Chart
 * Heatmap
 ************************************************************/


const CONFIG = {

  SHEET_NAME: 'DATA',
  ALARM_SHEET_NAME: 'ALARM_LOG',

  // ==============================
  // SENSOR THRESHOLDS
  // ==============================

  GAS_WARNING: 0.3,
  GAS_DANGER: 0.6,

  TEMP_NORMAL_MIN: 30,
  TEMP_WARNING: 45,
  TEMP_DANGER: 56,

  CO_WARNING: 0.3,
  CO_DANGER: 0.6,


  // ==============================
  // PREDICTION
  // ==============================

  PREDICTION_MINUTES: 2,

  HISTORY_POINTS: 60,

  DASHBOARD_POINTS: 120,


  // ==============================
  // ANOMALY
  // ==============================

  ANOMALY_POINTS: 30,

  ANOMALY_WARNING_Z: 2.0,

  ANOMALY_DANGER_Z: 3.0,


  // ==============================
  // REALTIME
  // ==============================

  REFRESH_SECONDS: 5,


  // ==============================
  // ALERT
  // ==============================

  ALERT_COOLDOWN_SECONDS: 300,


  // ==============================
  // AI
  // ==============================

  AI_MIN_INTERVAL_SECONDS: 60,

  GEMINI_MODEL: 'gemini-1.5-flash'

};


/************************************************************
 * WEB APP (doGet: ดึง Dashboard & รับค่าจาก ESP8266 ผ่าน GET)
 ************************************************************/

function doGet(e) {

  // ตรวจสอบว่าถ้ามี Sensor Parameter ส่งมาจาก ESP-12E ผ่าน HTTP GET ให้ประมวลผลเซนเซอร์
  if (e && e.parameter && (e.parameter.gas !== undefined || e.parameter.temp !== undefined || e.parameter.eco2 !== undefined)) {
    return handleSensorData(e.parameter);
  }

  // หากไม่มี Parameter ให้แสดงหน้าเว็บ Dashboard ตามปกติ
  return HtmlService
    .createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('EV SAFE GUARD')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/************************************************************
 * INCLUDE
 ************************************************************/

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


/************************************************************
 * ESP32 / ESP8266 POST API
 ************************************************************/

function doPost(e) {

  if (
    !e ||
    !e.postData ||
    !e.postData.contents
  ) {

    return jsonResponse({
      success: false,
      error: 'No POST data'
    });

  }

  try {
    const payload = JSON.parse(e.postData.contents);
    return handleSensorData(payload);
  } catch (err) {
    return jsonResponse({
      success: false,
      error: 'Invalid JSON payload'
    });
  }

}


/************************************************************
 * HANDLE SENSOR DATA (ระบบบันทึกและประมวลผลเซนเซอร์กลาง)
 ************************************************************/

function handleSensorData(params) {

  const lock = LockService.getScriptLock();

  try {

    lock.waitLock(10000);

    const gas = toNumber(params.gas);

    // รองรับชื่อพารามิเตอร์อุณหภูมิ (temp จาก ESP-12E หรือ temperature จาก payload เดิม)
    const temperature = toNumber(
      params.temperature ??
      params.temp
    );

    // รองรับชื่อพารามิเตอร์คาร์บอน/eCO2 (eco2 จาก ESP-12E หรือ co / CO จาก payload เดิม)
    const co = toNumber(
      params.co ??
      params.CO ??
      params.eco2
    );


    if (
      gas === null ||
      temperature === null ||
      co === null
    ) {

      return jsonResponse({
        success: false,
        error: 'Invalid sensor values'
      });

    }


    const timestamp =
      params.timestamp
        ? new Date(params.timestamp)
        : new Date();


    const sheet = getDataSheet();
    ensureDataHeaders(sheet);


    // ======================================
    // CURRENT SAFETY
    // ======================================

    const current = calculateSafety(
      gas,
      temperature,
      co
    );


    // ======================================
    // TEMPORARY ROW
    // ======================================

    sheet.appendRow([
      timestamp,
      gas,
      temperature,
      co,
      current.label,
      'กำลังวิเคราะห์...',
      current.risk,
      0,
      '→ STABLE',
      ''
    ]);


    const row = sheet.getLastRow();


    // ======================================
    // PREDICTION
    // ======================================

    const prediction = predictFuture(
      sheet,
      CONFIG.PREDICTION_MINUTES
    );


    // ======================================
    // ANOMALY
    // ======================================

    const anomaly = detectAnomaly(
      sheet,
      gas,
      temperature,
      co
    );


    // ======================================
    // FINAL RISK
    // ======================================

    const risk = calculateRisk(
      current,
      anomaly
    );


    // ======================================
    // TREND
    // ======================================

    const trend = determineTrend(sheet);


    // ======================================
    // AI ANALYSIS (ดึงค่าเดิมไว้ถ้า AI ไม่ได้รันรอบนี้)
    // ======================================

    let aiAnalysis = '';

    if (shouldRunAI()) {

      aiAnalysis = generateThaiAIAnalysis({
        gas: gas,
        temperature: temperature,
        co: co,
        current: current,
        prediction: prediction,
        anomaly: anomaly,
        risk: risk,
        trend: trend
      });

    } else {
      // หากยังไม่ถึงรอบรัน AI ให้ดึงผลวิเคราะห์ล่าสุดจากบรรทัดก่อนหน้ามาใช้ต่อ
      aiAnalysis = getLastAIAnalysis(sheet, row - 1);
    }


    // ======================================
    // UPDATE ROW
    // ======================================

    sheet
      .getRange(row, 6, 1, 5)
      .setValues([[
        prediction.label,
        risk.score,
        anomaly.score,
        trend,
        aiAnalysis
      ]]);


    // ======================================
    // ALARM ENGINE
    // ======================================

    processAlarm({
      timestamp: timestamp,
      gas: gas,
      temperature: temperature,
      co: co,
      current: current,
      prediction: prediction,
      anomaly: anomaly,
      risk: risk,
      trend: trend,
      aiAnalysis: aiAnalysis
    });


    SpreadsheetApp.flush();


    return jsonResponse({
      success: true,
      timestamp: timestamp.toISOString(),
      current: current,
      prediction: prediction,
      anomaly: anomaly,
      risk: risk,
      trend: trend,
      aiAnalysis: aiAnalysis
    });


  } catch (error) {

    return jsonResponse({
      success: false,
      error: error.toString()
    });

  } finally {

    try {
      lock.releaseLock();
    } catch (err) {}

  }

}


/************************************************************
 * DASHBOARD
 ************************************************************/

function getDashboardData() {

  const sheet = getDataSheet();
  ensureDataHeaders(sheet);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return emptyDashboard();
  }


  const startRow = Math.max(
    2,
    lastRow - CONFIG.DASHBOARD_POINTS + 1
  );


  const values = sheet
    .getRange(
      startRow,
      1,
      lastRow - startRow + 1,
      10
    )
    .getValues();


  const history = values.map(function(row) {
    return {
      timestamp: row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime(),
      time: formatTime(row[0]),
      gas: Number(row[1]) || 0,
      temperature: Number(row[2]) || 0,
      co: Number(row[3]) || 0,
      status: row[4] || '',
      prediction: row[5] || '',
      risk: Number(row[6]) || 0,
      anomaly: Number(row[7]) || 0,
      trend: row[8] || '',
      ai: row[9] || ''
    };
  });


  const last = history[history.length - 1];

  // ค้นหาข้อความ AI ล่าสุดที่มีอยู่ในประวัติ เพื่อป้องกันกรณีไม่มีข้อมูลส่งมา
  let latestAiAnalysis = last.ai || '';
  if (!latestAiAnalysis) {
    for (let i = history.length - 1; i >= 0; i--) {
      if (history[i].ai) {
        latestAiAnalysis = history[i].ai;
        break;
      }
    }
  }

  const current = calculateSafety(
    last.gas,
    last.temperature,
    last.co
  );

  const prediction = predictFuture(
    sheet,
    CONFIG.PREDICTION_MINUTES
  );

  const anomaly = detectAnomaly(
    sheet,
    last.gas,
    last.temperature,
    last.co
  );

  const risk = calculateRisk(
    current,
    anomaly
  );

  const trend = determineTrend(sheet);

  const alarms = getAlarmHistory(20);

  const heatmap = getHeatmapData();


  return {
    success: true,
    hasData: true,
    serverTime: Date.now(),
    current: {
      gas: last.gas,
      temperature: last.temperature,
      co: last.co
    },
    system: current,
    risk: risk,
    anomaly: anomaly,
    prediction: prediction,
    trend: trend,
    aiAnalysis: latestAiAnalysis,
    history: history,
    alarms: alarms,
    heatmap: heatmap,
    config: {
      refreshSeconds: CONFIG.REFRESH_SECONDS,
      predictionMinutes: CONFIG.PREDICTION_MINUTES
    }
  };

}


/************************************************************
 * HELPER: GET LAST AI ANALYSIS FROM SHEET
 ************************************************************/

function getLastAIAnalysis(sheet, fromRow) {
  if (fromRow < 2) return '';
  const start = Math.max(2, fromRow - 50);
  const values = sheet.getRange(start, 10, fromRow - start + 1, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] && String(values[i][0]).trim() !== '') {
      return String(values[i][0]);
    }
  }
  return '';
}


/************************************************************
 * HISTORICAL DATA
 ************************************************************/

function getHistoricalData(minutes) {

  minutes = Number(minutes) || 60;

  const sheet = getDataSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const totalRows = lastRow - 1;

  const values = sheet
    .getRange(2, 1, totalRows, 10)
    .getValues();

  const now = Date.now();
  const cutoff = now - minutes * 60 * 1000;

  const filtered = values.filter(function(row) {
    return (
      row[0] instanceof Date &&
      row[0].getTime() >= cutoff
    );
  });


  if (filtered.length <= 300) {
    return filtered.map(function(row) {
      return {
        timestamp: row[0].getTime(),
        time: formatTime(row[0]),
        gas: Number(row[1]) || 0,
        temperature: Number(row[2]) || 0,
        co: Number(row[3]) || 0,
        risk: Number(row[6]) || 0
      };
    });
  }


  // ======================================
  // AGGREGATE LARGE DATA
  // ======================================

  const bucketCount = 240;
  const bucketSize = (minutes * 60 * 1000) / bucketCount;
  const buckets = {};

  filtered.forEach(function(row) {
    const timestamp = row[0].getTime();
    const index = Math.floor((timestamp - cutoff) / bucketSize);

    if (!buckets[index]) {
      buckets[index] = {
        count: 0,
        gas: 0,
        temperature: 0,
        co: 0,
        risk: 0,
        timestamp: cutoff + index * bucketSize
      };
    }

    buckets[index].count++;
    buckets[index].gas += Number(row[1]) || 0;
    buckets[index].temperature += Number(row[2]) || 0;
    buckets[index].co += Number(row[3]) || 0;
    buckets[index].risk = Math.max(buckets[index].risk, Number(row[6]) || 0);
  });


  return Object.keys(buckets)
    .sort((a, b) => Number(a) - Number(b))
    .map(function(key) {
      const b = buckets[key];
      return {
        timestamp: b.timestamp,
        time: Utilities.formatDate(
          new Date(b.timestamp),
          Session.getScriptTimeZone(),
          'HH:mm'
        ),
        gas: round(b.gas / b.count, 2),
        temperature: round(b.temperature / b.count, 2),
        co: round(b.co / b.count, 2),
        risk: round(b.risk, 0)
      };
    });

}


/************************************************************
 * SAFETY ENGINE
 ************************************************************/

function calculateSafety(gas, temperature, co) {

  let level = 0;

  // GAS
  if (gas >= CONFIG.GAS_DANGER) {
    level = Math.max(level, 2);
  } else if (gas >= CONFIG.GAS_WARNING) {
    level = Math.max(level, 1);
  }

  // TEMPERATURE
  if (temperature >= CONFIG.TEMP_DANGER) {
    level = Math.max(level, 2);
  } else if (temperature >= CONFIG.TEMP_WARNING) {
    level = Math.max(level, 1);
  }

  // CO
  if (co >= CONFIG.CO_DANGER) {
    level = Math.max(level, 2);
  } else if (co >= CONFIG.CO_WARNING) {
    level = Math.max(level, 1);
  }

  const sensorRisk = calculateSensorRisk(gas, temperature, co);

  if (level === 2) {
    return {
      level: 2,
      label: '🔴 DANGER',
      text: 'อันตราย',
      risk: sensorRisk
    };
  }

  if (level === 1) {
    return {
      level: 1,
      label: '🟡 WARNING',
      text: 'เฝ้าระวัง',
      risk: sensorRisk
    };
  }

  return {
    level: 0,
    label: '🟢 NORMAL',
    text: 'ปกติ',
    risk: sensorRisk
  };

}


/************************************************************
 * RISK SCORE
 ************************************************************/

function calculateSensorRisk(gas, temperature, co) {

  const gasRisk = clamp((gas / CONFIG.GAS_DANGER) * 100, 0, 100);
  const coRisk = clamp((co / CONFIG.CO_DANGER) * 100, 0, 100);

  let tempRisk = 0;

  if (temperature < CONFIG.TEMP_WARNING) {
    tempRisk = 0;
  } else if (temperature < CONFIG.TEMP_DANGER) {
    tempRisk = 50 + ((temperature - CONFIG.TEMP_WARNING) / (CONFIG.TEMP_DANGER - CONFIG.TEMP_WARNING)) * 50;
  } else {
    tempRisk = 100;
  }

  return Math.round(Math.max(gasRisk, tempRisk, coRisk));

}


/************************************************************
 * FINAL RISK
 ************************************************************/

function calculateRisk(current, anomaly) {

  const score = Math.round(Math.max(current.risk, anomaly.score));
  let level = 0;

  if (score >= 70 || current.level === 2) {
    level = 2;
  } else if (score >= 30 || current.level === 1) {
    level = 1;
  }

  return {
    score: clamp(score, 0, 100),
    level: level,
    label: level === 2 ? '🔴 HIGH RISK' : level === 1 ? '🟡 MEDIUM RISK' : '🟢 LOW RISK'
  };

}


/************************************************************
 * ANOMALY DETECTION
 ************************************************************/

function detectAnomaly(sheet, gas, temperature, co) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 5) {
    return {
      score: 0,
      level: 0,
      label: '🟢 NORMAL',
      gasZ: 0,
      tempZ: 0,
      coZ: 0
    };
  }

  const startRow = Math.max(2, lastRow - CONFIG.ANOMALY_POINTS);
  const rows = sheet.getRange(startRow, 2, lastRow - startRow + 1, 3).getValues();

  const gasValues = rows.map(r => Number(r[0]));
  const tempValues = rows.map(r => Number(r[1]));
  const coValues = rows.map(r => Number(r[2]));

  const gasZ = zScore(gas, gasValues);
  const tempZ = zScore(temperature, tempValues);
  const coZ = zScore(co, coValues);

  const maxZ = Math.max(Math.abs(gasZ), Math.abs(tempZ), Math.abs(coZ));
  const score = clamp((maxZ / CONFIG.ANOMALY_DANGER_Z) * 100, 0, 100);

  let level = 0;
  if (maxZ >= CONFIG.ANOMALY_DANGER_Z) {
    level = 2;
  } else if (maxZ >= CONFIG.ANOMALY_WARNING_Z) {
    level = 1;
  }

  return {
    score: Math.round(score),
    level: level,
    label: level === 2 ? '🔴 ANOMALY' : level === 1 ? '🟡 UNUSUAL' : '🟢 NORMAL',
    gasZ: round(gasZ, 2),
    tempZ: round(tempZ, 2),
    coZ: round(coZ, 2)
  };

}


/************************************************************
 * Z SCORE
 ************************************************************/

function zScore(value, values) {

  if (values.length < 3) return 0;

  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  const variance = values.reduce((sum, x) => sum + Math.pow(x - mean, 2), 0) / values.length;
  const std = Math.sqrt(variance);

  if (std === 0) return 0;

  return (value - mean) / std;

}


/************************************************************
 * 2-MINUTE PREDICTION
 ************************************************************/

function predictFuture(sheet, minutes) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 5) {
    return {
      available: false,
      gas: 0,
      temperature: 0,
      co: 0,
      label: '⏳ กำลังเก็บข้อมูล',
      confidence: 0,
      trend: 'WAITING'
    };
  }

  const startRow = Math.max(2, lastRow - CONFIG.HISTORY_POINTS + 1);
  const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, 4)
    .getValues()
    .filter(r => r[0] instanceof Date);

  if (rows.length < 3) {
    return {
      available: false,
      gas: 0,
      temperature: 0,
      co: 0,
      label: '⏳ กำลังเก็บข้อมูล',
      confidence: 0,
      trend: 'WAITING'
    };
  }

  const latestTime = rows[rows.length - 1][0];
  const horizon = minutes * 60 * 1000;

  const gas = regression(rows, 1, latestTime, horizon);
  const temp = regression(rows, 2, latestTime, horizon);
  const co = regression(rows, 3, latestTime, horizon);

  const predictedSafety = calculateSafety(gas.value, temp.value, co.value);
  const confidence = Math.round(((gas.r2 + temp.r2 + co.r2) / 3) * 100);

  return {
    available: true,
    gas: round(Math.max(0, gas.value), 2),
    temperature: round(Math.max(0, temp.value), 2),
    co: round(Math.max(0, co.value), 2),
    label: predictedSafety.label,
    status: predictedSafety,
    confidence: confidence,
    horizonMinutes: minutes
  };

}


/************************************************************
 * LINEAR REGRESSION
 ************************************************************/

function regression(rows, columnIndex, latestTime, horizon) {

  const x = [];
  const y = [];

  rows.forEach(function(row) {
    x.push((row[0].getTime() - latestTime.getTime()) / 1000);
    y.push(Number(row[columnIndex]));
  });

  const n = x.length;
  let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;

  for (let i = 0; i < n; i++) {
    sumX += x[i];
    sumY += y[i];
    sumXY += x[i] * y[i];
    sumX2 += x[i] * x[i];
  }

  const denominator = n * sumX2 - sumX * sumX;

  if (denominator === 0) {
    return { value: y[y.length - 1], r2: 0 };
  }

  const slope = (n * sumXY - sumX * sumY) / denominator;
  const intercept = (sumY - slope * sumX) / n;

  const futureSeconds = horizon / 1000;
  const value = intercept + slope * futureSeconds;

  const meanY = sumY / n;
  let ssTot = 0, ssRes = 0;

  for (let i = 0; i < n; i++) {
    const predicted = intercept + slope * x[i];
    ssTot += Math.pow(y[i] - meanY, 2);
    ssRes += Math.pow(y[i] - predicted, 2);
  }

  const r2 = ssTot === 0 ? 0 : clamp(1 - ssRes / ssTot, 0, 1);

  return { value: value, r2: r2 };

}


/************************************************************
 * TREND
 ************************************************************/

function determineTrend(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 4) return '→ STABLE';

  const startRow = Math.max(2, lastRow - 4);
  const rows = sheet.getRange(startRow, 2, lastRow - startRow + 1, 3).getValues();

  const first = rows[0];
  const last = rows[rows.length - 1];

  const gasChange = Number(last[0]) - Number(first[0]);
  const tempChange = Number(last[1]) - Number(first[1]);
  const coChange = Number(last[2]) - Number(first[2]);

  const score = gasChange + tempChange + coChange;

  if (score > 0.05) return '↗ INCREASING';
  if (score < -0.05) return '↘ DECREASING';

  return '→ STABLE';

}


/************************************************************
 * ALARM ENGINE
 ************************************************************/

function processAlarm(data) {

  const properties = PropertiesService.getScriptProperties();
  const previousLevel = Number(properties.getProperty('CURRENT_ALARM_LEVEL') || 0);

  const currentLevel = Math.max(
    data.current.level,
    data.risk.level,
    data.anomaly.level,
    data.prediction.status ? data.prediction.status.level : 0
  );

  let source = 'CURRENT';

  if (data.prediction.status && data.prediction.status.level > data.current.level) {
    source = 'AI PREDICTION';
  }

  if (data.anomaly.level > data.current.level) {
    source = 'ANOMALY';
  }


  // NORMAL -> WARNING/DANGER
  if (currentLevel > previousLevel) {

    writeAlarm({
      timestamp: data.timestamp,
      level: currentLevel,
      source: source,
      gas: data.gas,
      temperature: data.temperature,
      co: data.co,
      risk: data.risk.score,
      anomaly: data.anomaly.score,
      prediction: data.prediction.label,
      aiAnalysis: data.aiAnalysis
    });

    sendLineAlert({
      type: 'ALERT',
      level: currentLevel,
      source: source,
      gas: data.gas,
      temperature: data.temperature,
      co: data.co,
      risk: data.risk.score,
      anomaly: data.anomaly.score,
      prediction: data.prediction.label,
      analysis: data.aiAnalysis
    });

  }
  // DANGER CONTINUES
  else if (currentLevel === 2) {

    const lastSent = Number(properties.getProperty('LAST_ALERT_TIME') || 0);
    const cooldown = CONFIG.ALERT_COOLDOWN_SECONDS * 1000;

    if (Date.now() - lastSent > cooldown) {

      writeAlarm({
        timestamp: data.timestamp,
        level: 2,
        source: source,
        gas: data.gas,
        temperature: data.temperature,
        co: data.co,
        risk: data.risk.score,
        anomaly: data.anomaly.score,
        prediction: data.prediction.label,
        aiAnalysis: data.aiAnalysis
      });

      sendLineAlert({
        type: 'REPEAT_ALERT',
        level: 2,
        source: source,
        gas: data.gas,
        temperature: data.temperature,
        co: data.co,
        risk: data.risk.score,
        anomaly: data.anomaly.score,
        prediction: data.prediction.label,
        analysis: data.aiAnalysis
      });

    }

  }

  // RECOVERY
  if (previousLevel > 0 && currentLevel === 0) {

    writeAlarm({
      timestamp: data.timestamp,
      level: 0,
      source: 'RECOVERY',
      gas: data.gas,
      temperature: data.temperature,
      co: data.co,
      risk: data.risk.score,
      anomaly: data.anomaly.score,
      prediction: data.prediction.label,
      aiAnalysis: data.aiAnalysis
    });

    sendLineAlert({
      type: 'RECOVERY',
      level: 0,
      source: 'RECOVERY',
      gas: data.gas,
      temperature: data.temperature,
      co: data.co,
      risk: data.risk.score,
      anomaly: data.anomaly.score,
      prediction: data.prediction.label,
      analysis: data.aiAnalysis
    });

  }

  properties.setProperty('CURRENT_ALARM_LEVEL', String(currentLevel));

}


/************************************************************
 * WRITE ALARM
 ************************************************************/

function writeAlarm(data) {

  const sheet = getAlarmSheet();
  ensureAlarmHeaders(sheet);

  sheet.appendRow([
    data.timestamp,
    getAlarmLabel(data.level),
    data.source,
    data.gas,
    data.temperature,
    data.co,
    data.risk,
    data.anomaly,
    data.prediction,
    data.aiAnalysis
  ]);

}


/************************************************************
 * ALARM HISTORY
 ************************************************************/

function getAlarmHistory(limit) {

  const sheet = getAlarmSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const start = Math.max(2, lastRow - limit + 1);
  const values = sheet.getRange(start, 1, lastRow - start + 1, 10).getValues();

  return values.reverse().map(function(row) {
    return {
      timestamp: row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime(),
      time: formatDateTime(row[0]),
      level: row[1],
      source: row[2],
      gas: Number(row[3]) || 0,
      temperature: Number(row[4]) || 0,
      co: Number(row[5]) || 0,
      risk: Number(row[6]) || 0,
      anomaly: Number(row[7]) || 0,
      prediction: row[8] || '',
      analysis: row[9] || ''
    };
  });

}


/************************************************************
 * HEATMAP
 ************************************************************/

function getHeatmapData() {

  const sheet = getAlarmSheet();
  const lastRow = sheet.getLastRow();
  const result = [];

  for (let day = 0; day < 7; day++) {
    result[day] = new Array(24).fill(0);
  }

  if (lastRow < 2) return result;

  const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000;
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  values.forEach(function(row) {
    if (!(row[0] instanceof Date)) return;

    const timestamp = row[0].getTime();
    if (timestamp < cutoff) return;

    const date = row[0];
    const day = date.getDay();
    const hour = date.getHours();
    const level = alarmLabelToLevel(row[1]);

    result[day][hour] = Math.max(result[day][hour], level);
  });

  return result;

}


/************************************************************
 * LINE NOTIFICATION (ปรับแจ้งเตือนเฉพาะ สีเหลือง/แดง)
 ************************************************************/

function sendLineAlert(data) {

  // ตรวจเช็กระดับความปลอดภัย หากไม่ใช่ สีเหลือง (1) หรือ สีแดง (2) หรือเป็น RECOVERY ให้ข้ามการส่ง LINE
  if (data.type === 'RECOVERY' || (data.level !== 1 && data.level !== 2)) {
    return { success: true, skipped: true, message: 'Skipped notification for Normal/Recovery state' };
  }

  const properties = PropertiesService.getScriptProperties();
  const token = properties.getProperty('LINE_TOKEN');
  const userId = properties.getProperty('LINE_USER_ID');

  if (!token || !userId) {
    return { success: false, error: 'LINE configuration missing' };
  }

  let headerBg = '#F39C12'; // สีเหลือง default
  let headerTitle = '⚠️ EV SAFE GUARD : WARNING';
  let statusText = 'เฝ้าระวังความเสี่ยง';

  if (data.level === 2) {
    headerBg = '#E74C3C'; // สีแดง
    headerTitle = '🚨 EV SAFE GUARD : DANGER';
    statusText = 'อันตรายระดับสูง!';
  } else if (data.level === 1) {
    headerBg = '#F39C12'; // สีเหลือง
    headerTitle = '⚠️ EV SAFE GUARD : WARNING';
    statusText = 'เฝ้าระวังความเสี่ยง';
  }

  // ใช้ Flex Message แบบการ์ด
  const flexPayload = {
    to: userId,
    messages: [{
      type: 'flex',
      altText: `${headerTitle} - ${statusText}`,
      contents: {
        type: 'bubble',
        header: {
          type: 'box',
          layout: 'vertical',
          backgroundColor: headerBg,
          paddingAll: 'lg',
          contents: [
            { type: 'text', text: headerTitle, weight: 'bold', color: '#FFFFFF', size: 'md' },
            { type: 'text', text: `สถานะ: ${statusText}`, color: '#FFFFFF', size: 'xs', margin: 'xs' },
            { type: 'text', text: `แหล่งที่มา: ${data.source || 'SYSTEM'}`, color: '#FFFFFF', size: 'xxs', margin: 'xs' }
          ]
        },
        body: {
          type: 'box',
          layout: 'vertical',
          spacing: 'md',
          contents: [
            {
              type: 'box',
              layout: 'horizontal',
              contents: [
                { type: 'text', text: 'Gas', size: 'sm', color: '#888888', flex: 2 },
                { type: 'text', text: `${data.gas}`, size: 'sm', weight: 'bold', align: 'end', flex: 3 }
              ]
            },
            {
              type: 'box',
              layout: 'horizontal',
              contents: [
                { type: 'text', text: 'อุณหภูมิ', size: 'sm', color: '#888888', flex: 2 },
                { type: 'text', text: `${data.temperature} °C`, size: 'sm', weight: 'bold', align: 'end', flex: 3 }
              ]
            },
            {
              type: 'box',
              layout: 'horizontal',
              contents: [
                { type: 'text', text: 'CO', size: 'sm', color: '#888888', flex: 2 },
                { type: 'text', text: `${data.co}`, size: 'sm', weight: 'bold', align: 'end', flex: 3 }
              ]
            },
            { type: 'separator', margin: 'md' },
            {
              type: 'box',
              layout: 'horizontal',
              margin: 'md',
              contents: [
                { type: 'text', text: 'Risk Score', size: 'sm', color: '#555555', flex: 2 },
                { type: 'text', text: `${data.risk}/100`, size: 'sm', weight: 'bold', color: data.risk > 50 ? '#E74C3C' : '#27AE60', align: 'end', flex: 3 }
              ]
            },
            {
              type: 'box',
              layout: 'horizontal',
              contents: [
                { type: 'text', text: 'Anomaly Score', size: 'sm', color: '#555555', flex: 2 },
                { type: 'text', text: `${data.anomaly}/100`, size: 'sm', weight: 'bold', color: data.anomaly > 50 ? '#E74C3C' : '#27AE60', align: 'end', flex: 3 }
              ]
            },
            {
              type: 'box',
              layout: 'horizontal',
              contents: [
                { type: 'text', text: 'AI Prediction', size: 'sm', color: '#555555', flex: 2 },
                { type: 'text', text: `${data.prediction}`, size: 'sm', weight: 'bold', align: 'end', flex: 3 }
              ]
            },
            ...(data.analysis ? [
              { type: 'separator', margin: 'md' },
              { type: 'text', text: '🤖 AI วิเคราะห์:', size: 'xs', weight: 'bold', color: '#333333', margin: 'md' },
              { type: 'text', text: data.analysis, size: 'xxs', color: '#666666', wrap: true }
            ] : [])
          ]
        }
      }
    }]
  };

  const url = 'https://api.line.me/v2/bot/message/push';

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(flexPayload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();

  if (code === 200) {
    properties.setProperty('LAST_ALERT_TIME', String(Date.now()));
  }

  return {
    success: code === 200,
    code: code,
    response: response.getContentText()
  };

}


/************************************************************
 * AI THAI ANALYSIS (GEMINI)
 ************************************************************/

function generateThaiAIAnalysis(context) {

  const properties = PropertiesService.getScriptProperties();
  const apiKey = properties.getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    return generateRuleThaiAnalysis(context);
  }

  const model = properties.getProperty('GEMINI_MODEL') || CONFIG.GEMINI_MODEL;
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(apiKey);

  const prompt = `
คุณคือ AI วิเคราะห์ความปลอดภัยของระบบ EV SAFE GUARD

วิเคราะห์ข้อมูล Sensor ต่อไปนี้เป็นภาษาไทย

Gas: ${context.gas}
Temperature: ${context.temperature} °C
CO: ${context.co}

Current Status:
${context.current.label}

Risk Score:
${context.risk.score}/100

Anomaly:
${context.anomaly.label}
Score ${context.anomaly.score}

Prediction +2 minutes:
${context.prediction.label}

Predicted Gas:
${context.prediction.gas}

Predicted Temperature:
${context.prediction.temperature}

Predicted CO:
${context.prediction.co}

Trend:
${context.trend}

ตอบสั้น กระชับ ไม่เกิน 500 ตัวอักษร

รูปแบบ:

สถานการณ์:
...

แนวโน้ม:
...

ความเสี่ยง:
...

คำแนะนำ:
...

ห้ามอ้างว่าข้อมูลนี้เป็นการรับรองความปลอดภัย
`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      return generateRuleThaiAnalysis(context);
    }

    const json = JSON.parse(response.getContentText());
    const text = json?.candidates?.[0]?.content?.parts?.[0]?.text;

    if (text) return text.trim();

    return generateRuleThaiAnalysis(context);

  } catch (error) {
    console.log('Gemini error: ' + error);
    return generateRuleThaiAnalysis(context);
  }

}


/************************************************************
 * RULE BASED THAI AI FALLBACK
 ************************************************************/

function generateRuleThaiAnalysis(context) {

  const situations = [];

  if (context.current.level === 2) {
    situations.push('พบค่าที่อยู่ในระดับอันตราย');
  } else if (context.current.level === 1) {
    situations.push('พบค่าที่ต้องเฝ้าระวัง');
  } else {
    situations.push('ค่าปัจจุบันยังอยู่ในระดับปกติ');
  }

  if (context.anomaly.level === 2) {
    situations.push('ตรวจพบความผิดปกติของรูปแบบข้อมูล');
  }

  let trendText = 'แนวโน้มค่อนข้างคงที่';
  if (context.trend === '↗ INCREASING') trendText = 'ค่ามีแนวโน้มเพิ่มขึ้น';
  else if (context.trend === '↘ DECREASING') trendText = 'ค่ามีแนวโน้มลดลง';

  let riskText = 'ความเสี่ยงอยู่ในระดับต่ำ';
  if (context.risk.score >= 70) {
    riskText = 'ความเสี่ยงอยู่ในระดับสูง ควรตรวจสอบระบบทันที';
  } else if (context.risk.score >= 30) {
    riskText = 'ความเสี่ยงอยู่ในระดับปานกลาง ควรเฝ้าระวัง';
  }

  let advice = 'ติดตามค่า Sensor ต่อเนื่อง';
  if (context.risk.score >= 70) {
    advice = 'ตรวจสอบแหล่งกำเนิดค่า Sensor และดำเนินการตามขั้นตอนความปลอดภัยของหน่วยงาน';
  } else if (context.prediction.status && context.prediction.status.level >= 2) {
    advice = 'ระบบคาดการณ์ว่าอาจเข้าสู่ระดับอันตรายใน 2 นาที ควรตรวจสอบล่วงหน้า';
  }

  return (
    'สถานการณ์: ' + situations.join(' ') +
    '\nแนวโน้ม: ' + trendText +
    '\nความเสี่ยง: ' + riskText +
    '\nคำแนะนำ: ' + advice
  );

}


/************************************************************
 * AI THROTTLE
 ************************************************************/

function shouldRunAI() {

  const cache = CacheService.getScriptCache();
  const last = Number(cache.get('LAST_AI_TIME') || 0);
  const now = Date.now();

  if (now - last < CONFIG.AI_MIN_INTERVAL_SECONDS * 1000) {
    return false;
  }

  cache.put('LAST_AI_TIME', String(now), CONFIG.AI_MIN_INTERVAL_SECONDS);

  return true;

}


/************************************************************
 * SHEETS MANAGEMENT
 ************************************************************/

function getDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  return sheet;
}

function getAlarmSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.ALARM_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.ALARM_SHEET_NAME);
  }
  return sheet;
}


/************************************************************
 * HEADERS
 ************************************************************/

function ensureDataHeaders(sheet) {

  const headers = [
    'Date/Time',
    'Gas',
    'อุณหภูมิ',
    'คาร์บอนมอนอกไซด์ (CO)',
    'Status',
    'AI Prediction',
    'Risk Score',
    'Anomaly Score',
    'Trend',
    'AI Analysis'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd HH:mm:ss');
  sheet.getRange('B:D').setNumberFormat('0.00');

}

function ensureAlarmHeaders(sheet) {

  const headers = [
    'Date/Time',
    'Level',
    'Source',
    'Gas',
    'Temperature',
    'CO',
    'Risk Score',
    'Anomaly Score',
    'Prediction',
    'AI Analysis'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd HH:mm:ss');

}


/************************************************************
 * TEST FUNCTIONS
 ************************************************************/

function testInsertData() {

  const gas = Number((Math.random() * 0.8).toFixed(2));
  const temperature = Number((35 + Math.random() * 25).toFixed(1));
  const co = Number((Math.random() * 0.8).toFixed(2));

  const payload = {
    gas: gas,
    temperature: temperature,
    co: co
  };

  const result = handleSensorData(payload);
  console.log(result.getContent());

}

function testLINE() {

  return sendLineAlert({
    type: 'TEST',
    level: 1,
    source: 'SYSTEM TEST',
    gas: 0.25,
    temperature: 48,
    co: 0.20,
    risk: 50,
    anomaly: 20,
    prediction: '🟡 WARNING',
    analysis: 'นี่คือข้อความทดสอบระบบแจ้งเตือน EV SAFE GUARD'
  });

}

function testGemini() {

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const model = props.getProperty('GEMINI_MODEL') || 'gemini-1.5-flash';

  if (!apiKey) {
    throw new Error('ไม่พบ GEMINI_API_KEY ใน Script Properties');
  }

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    model + ':generateContent';

  const payload = {
    contents: [{ parts: [{ text: 'ตอบเป็นภาษาไทยว่า: ระบบ Gemini เชื่อมต่อสำเร็จหรือไม่' }] }]
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  Logger.log(JSON.stringify(result, null, 2));

  if (result.candidates && result.candidates[0]) {
    const text = result.candidates[0].content.parts[0].text;
    Logger.log('Gemini ตอบ: ' + text);
  } else {
    Logger.log('เกิดข้อผิดพลาด: ' + response.getContentText());
  }

}


/************************************************************
 * UTILITIES
 ************************************************************/

function emptyDashboard() {

  return {
    success: true,
    hasData: false,
    current: { gas: 0, temperature: 0, co: 0 },
    system: { level: 0, label: '🟢 NORMAL', risk: 0 },
    risk: { score: 0, level: 0, label: '🟢 LOW RISK' },
    anomaly: { score: 0, level: 0, label: '🟢 NORMAL' },
    prediction: { available: false, label: '⏳ รอข้อมูล' },
    trend: '→ STABLE',
    history: [],
    alarms: [],
    heatmap: []
  };

}

function getAlarmLabel(level) {
  if (level === 2) return '🔴 DANGER';
  if (level === 1) return '🟡 WARNING';
  return '🟢 RECOVERY';
}

function alarmLabelToLevel(label) {
  if (String(label).includes('DANGER')) return 2;
  if (String(label).includes('WARNING')) return 1;
  return 0;
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  const number = Number(value);
  return isFinite(number) ? number : null;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function round(value, digits) {
  const factor = Math.pow(10, digits);
  return Math.round(value * factor) / factor;
}

function formatTime(value) {
  if (!(value instanceof Date)) return '';
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm:ss');
}

function formatDateTime(value) {
  if (!(value instanceof Date)) return '';
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}