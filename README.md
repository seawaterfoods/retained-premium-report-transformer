# 自留保費統計表報表轉換系統

> 將多個產險公司的「自留保費統計表」Excel 來源檔，自動彙整產生報表。

**Java 17** · **Spring Boot 3.5.0** · **Apache POI 5.3.0** · **Maven** · **Docker**

## 📖 文件

完整文件請參閱 [docs/README.md](docs/README.md)。

## 🚀 快速開始

```bash
# 1. 編譯
build.bat

# 2. 準備資料：將來源 Excel 放入 import/{年度}/Q{季度}/
#    例如: import/115/Q1/29_115(01-03)_自留保費統計表.xlsx

# 3. 設定：編輯 config/application.yml (設定 process-year)

# 4. 執行
run.bat

# 5. 查看輸出
ls output/115Q1/
```

## 🐳 Docker

```bash
docker-compose build
docker-compose up
```

## 專案結構

```
├── config/          外部設定檔 (使用者可修改)
├── import/          來源資料 (按年度/季度分目錄)
│   └── {year}/Q{quarter}/   例如 115/Q1/
├── output/          產出報表 (按年度季度分目錄)
│   └── {year}Q{quarter}/    例如 115Q1/
├── src/main/java/com/insurance/retainedpremium/
│   ├── config/      配置
│   ├── model/       資料模型
│   ├── reader/      Excel 讀取與驗證
│   ├── writer/      Excel 寫入
│   ├── service/     業務邏輯編排
│   ├── constant/    常數定義
│   └── exception/   例外處理
├── docs/            完整文件
├── build.bat        編譯腳本
├── run.bat          執行腳本
└── pom.xml          Maven 建置設定
```

## 功能摘要

| 功能 | 說明 |
|------|------|
| 多檔案讀取 | 同時匯入多家公司的來源 Excel |
| 自動險種歸類 | 33 種險種自動歸類為 16 大類 |
| 季度判斷 | 依檔名自動判定季度 (Q1–Q4) |
| 公式保留 | 輸出報表完整保留 Excel 公式 |
| 動態公司顯示 | 僅顯示有資料的公司，隱藏空列 |
| 去年同期比較 | 自動從 output/{year-1}Q{quarter}/ 讀取去年報表 |
| 容器化部署 | 支援 Docker 一鍵執行 |
