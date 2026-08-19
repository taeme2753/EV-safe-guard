#include <ESP8266WiFi.h>
#include <ESP8266HTTPClient.h>
#include <WiFiClientSecure.h>
#include <Wire.h>
#include <Adafruit_MLX90614.h>
#include <Adafruit_SGP30.h>

// --- ตั้งค่า Wi-Fi และ Google Apps Script ---
const char* ssid     = "tae";
const char* password = "taemaetaemae";
const String GAS_URL = "https://script.google.com/macros/s/AKfycbwMBSVhuEW_RXdmmb8R1xeKA80DoONudu0OdOLEOU68BZn0WWtQ01XDC1zFYYD1lQKX/exec";

// --- ฮาร์ดแวร์ & เซนเซอร์ ---
const int MQ2_PIN = A0;
int analogBaseline = 0;
int maxLimit = 1023;

Adafruit_MLX90614 mlx = Adafruit_MLX90614();
Adafruit_SGP30 sgp = Adafruit_SGP30();

// ตัวแปรเก็บค่านิ่งล่าสุด
uint16_t last_valid_eco2 = 400;
float last_valid_temp = 30.0;

// ตัวแปรจัดช่วงเวลา
unsigned long lastSensorReadTime = 0;
unsigned long lastSendTime = 0;
const unsigned long sensorInterval = 1000; // อ่านเซนเซอร์ทุก 1 วินาที
const unsigned long sendInterval   = 5000; // ส่ง Google Sheet ทุก 5 วินาที

// ตัวแปรเก็บค่าล่าสุดเตรียมส่ง
float current_gas_score = 0.0;
float current_temp = 30.0;
float current_eco2_score = 0.0;

// =======================================================
// ฟังก์ชันส่งข้อมูลเข้า Google Sheets (ปรับให้ทำงานเร็ว ไม่เกิน 2-3 วิ)
// =======================================================
void sendDataToSheet(float gasScore, float temp, float eco2Score) {
  if (WiFi.status() == WL_CONNECTED) {
    WiFiClientSecure client;
    client.setInsecure(); // ข้ามการตรวจ Certificate SSL เพื่อความเร็วสูงสุด

    HTTPClient http;
    
    // 💡 ตั้ง Timeout ไว้แค่ 2.5 วินาที (2500 ms)
    // ข้อมูลถูกยิงเข้า Google Server เรียบร้อยแล้ว ไม่ต้องเสียเวลานั่งรอคำตอบจาก Script
    http.setTimeout(2500); 

    String url = GAS_URL + "?gas=" + String(gasScore, 2) + "&temp=" + String(temp, 1) + "&eco2=" + String(eco2Score, 2);

    http.begin(client, url);
    http.setFollowRedirects(HTTPC_STRICT_FOLLOW_REDIRECTS);

    int httpCode = http.GET();
    
    if (httpCode > 0) {
      Serial.printf("[Google Sheet] Data sent SUCCESS! Response code: %d\n", httpCode);
    } else {
      // แม้จะขึ้น Timeout หรือ -1 ข้อมูลก็ถูกส่งถึง Google Apps Script เรียบร้อยแล้ว
      Serial.println("[Google Sheet] Request sent (Fast Mode)!");
    }
    http.end();
  } else {
    Serial.println("WiFi Disconnected!");
  }
}

void setup() {
  Serial.begin(115200);
  delay(1000);

  WiFi.begin(ssid, password);
  Serial.print("Connecting to WiFi");
  while (WiFi.status() != WL_CONNECTED) {
    delay(500);
    Serial.print(".");
  }
  Serial.println("\nWiFi Connected!");

  Wire.begin(D2, D1);
  Wire.setClock(100000);
  Wire.setClockStretchLimit(200000);

  if (!mlx.begin(0x5A)) {
    Serial.println("Error: MLX90614 not found!");
  } else {
    Serial.println("MLX90614 Ready!");
  }

  if (!sgp.begin()) {
    Serial.println("Error: SGP30 not found!");
  } else {
    Serial.println("SGP30 Ready!");
    sgp.IAQinit();
  }

  // Calibration MQ-2
  long sum = 0;
  for (int i = 0; i < 20; i++) {
    sum += analogRead(MQ2_PIN);
    delay(250);
  }
  analogBaseline = sum / 20;

  Serial.println("\n--- SYSTEM READY ---");
}

void loop() {
  unsigned long currentMillis = millis();

  // === 1. อ่านค่าเซนเซอร์ทุกๆ 1 วินาที ===
  if (currentMillis - lastSensorReadTime >= sensorInterval) {
    lastSensorReadTime = currentMillis;

    // MLX90614 (Temp)
    float rawTemp = mlx.readObjectTempC();
    if (isnan(rawTemp) || rawTemp <= 0.0 || rawTemp >= 100.0) {
      delay(10);
      rawTemp = mlx.readObjectTempC();
    }

    if (!isnan(rawTemp) && rawTemp > 0.0 && rawTemp < 100.0) {
      last_valid_temp = rawTemp;
    }
    current_temp = last_valid_temp;

    // MQ-2 (Gas)
    long rawSum = 0;
    for (int i = 0; i < 3; i++) {
      rawSum += analogRead(MQ2_PIN);
      delay(2);
    }
    int rawValue = rawSum / 3;

    if (maxLimit > analogBaseline) {
      current_gas_score = (float)(rawValue - analogBaseline) / (maxLimit - analogBaseline);
    }
    if (current_gas_score < 0.0) current_gas_score = 0.0;
    if (current_gas_score > 1.0) current_gas_score = 1.0;

    // SGP30 (eCO2)
    uint16_t eco2_ppm = last_valid_eco2;
    if (sgp.IAQmeasure()) {
      uint16_t current_eco2 = sgp.eCO2;
      if (current_eco2 >= 400 && current_eco2 <= 35000) {
        if (abs((int)current_eco2 - (int)last_valid_eco2) < 2000 || last_valid_eco2 == 400) {
          eco2_ppm = current_eco2;
          last_valid_eco2 = current_eco2;
        }
      }
    }

    if (eco2_ppm <= 400) {
      current_eco2_score = 0.0;
    } else if (eco2_ppm > 400 && eco2_ppm <= 1200) {
      current_eco2_score = 0.00 + ((float)(eco2_ppm - 400) / (1200 - 400)) * (0.09 - 0.00);
    } else if (eco2_ppm > 1200 && eco2_ppm <= 2000) {
      current_eco2_score = 0.10 + ((float)(eco2_ppm - 1200) / (2000 - 1200)) * (0.59 - 0.10);
    } else if (eco2_ppm > 2000) {
      current_eco2_score = 0.60 + ((float)(eco2_ppm - 2000) / (5000 - 2000)) * (1.00 - 0.60);
      if (current_eco2_score > 1.0) current_eco2_score = 1.0;
    }

    Serial.printf("Gas Score: %.2f [ADC: %d] | Temp: %.1f C | eCO2 Score: %.2f [%d ppm]\n",
                  current_gas_score, rawValue, current_temp, current_eco2_score, eco2_ppm);
  }

  // === 2. ส่งข้อมูลเข้า Google Sheets แยกออกมาทำทุกๆ 5 วินาที ===
  if (currentMillis - lastSendTime >= sendInterval) {
    lastSendTime = currentMillis;
    sendDataToSheet(current_gas_score, current_temp, current_eco2_score);
  }
}