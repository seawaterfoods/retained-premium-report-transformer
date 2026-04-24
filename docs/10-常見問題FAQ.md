← [回到索引](README.md)

# 第十章：常見問題 FAQ

---

### Q: 程式執行後沒有產出檔案？

檢查 log 輸出。常見原因：
1. `import/` 目錄中沒有 `.xlsx` 檔案
2. 檔名格式不符規定
3. 來源資料有誤導致中止

### Q: 輸出報表的公式沒有計算結果？

這是正常的。程式設定了 `setForceFormulaRecalculation(true)`，用 Excel 開啟檔案時會自動重新計算所有公式。

### Q: 可以同時處理不同年度嗎？

不可以。系統會驗證所有來源檔的年度必須一致，不同年度會中止處理。

### Q: 新增一家公司怎麼辦？

只要在 `import/` 目錄放入該公司的來源 `.xlsx` 檔案即可。系統會自動從來源資料中讀取公司資訊，程式化產生報表時會自動包含所有有資料的公司。

### Q: Docker 環境如何看 log？

```bash
docker-compose logs -f report-transformer
```

### Q: 如何修改日誌等級？

方式一：修改 `config/application.yml` 中的 logging 設定

方式二：透過環境變數
```bash
LOGGING_LEVEL_COM_INSURANCE_RETAINEDPREMIUM=DEBUG
```
