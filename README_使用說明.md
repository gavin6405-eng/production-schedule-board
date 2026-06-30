# 生產排程反推看板 v2

## 功能
- 上傳 Excel / CSV 生產排程
- 依組立地點反推建議入庫日
- 依 Category 或人工標準工期反推建議發料日
- 納入最晚到料日計算實際可開工日
- 自動判定正常、緩衝不足、已逾期、排程需變更
- 匯出 Excel 分析結果

## Streamlit 部署設定
- Repository：你的 GitHub 儲存庫
- Branch：main
- Main file path：app.py

## 必要欄位
- 組立地點
- 客戶入庫日

## 建議欄位
- 製令
- 客戶
- P/N
- Type
- Category
- 組立進度
- 備註
- 發料日
- 入庫日
- 最晚到料日
- 標準工期
