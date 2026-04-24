# 自留保費統計表報表轉換系統

> 將多個產險公司的「自留保費統計表」Excel 來源檔，自動彙整填入輸出模板，產生季度報表。

**Java 17** · **Spring Boot 3.5.0** · **Apache POI 5.3.0** · **Maven** · **Docker**

## 📖 文件

完整文件請參閱 [docs/README.md](docs/README.md)。

## 🚀 快速開始

```bash
# 1. 編譯
build.bat

# 2. 準備資料：將來源 Excel 放入 import/
#    將模板檔放到 import/template.xlsx
#    將去年報表放到 import/lastyear/

# 3. 設定：編輯 config/application.yml

# 4. 執行
run.bat

# 5. 查看輸出
ls output/
```

## 🐳 Docker

```bash
docker-compose build
docker-compose up
```

## 專案結構

```
├── config/          外部設定檔 (使用者可修改)
├── import/          來源資料 (使用者放入)
├── output/          產出報表
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
| 去年同期比較 | 自動讀取去年報表填入對照欄 |
| 容器化部署 | 支援 Docker 一鍵執行 |
