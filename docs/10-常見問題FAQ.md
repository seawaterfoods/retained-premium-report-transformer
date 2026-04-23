← [回到索引](README.md)

# 第十章：常見問題 FAQ

---

### Q: 程式執行後沒有產出檔案？

檢查 log 輸出。常見原因：
1. `input/` 目錄中沒有 `.xlsx` 檔案
2. 檔名格式不符規定
3. `templates/template.xlsx` 不存在
4. 來源資料有誤導致中止

### Q: 輸出報表的公式沒有計算結果？

這是正常的。程式設定了 `setForceFormulaRecalculation(true)`，用 Excel 開啟檔案時會自動重新計算所有公式。

### Q: 可以同時處理不同年度嗎？

不可以。系統會驗證所有來源檔的年度必須一致，不同年度會中止處理。

### Q: 新增一家公司（不在模板的 19 家中）怎麼辦？

目前需修改模板 Excel，在各季度區塊中新增該公司的列。程式會自動讀取模板中的公司代號進行匹配。

### Q: Docker 環境如何看 log？

```bash
docker-compose logs -f report-transformer
```

### Q: 如何修改日誌等級？

方式一：修改 `application.yml`
```yaml
logging:
  level:
    com.example.retainedpremium: DEBUG
```

方式二：透過環境變數
```bash
LOGGING_LEVEL_COM_EXAMPLE_RETAINEDPREMIUM=DEBUG
```
