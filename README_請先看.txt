100大股票水溫儀表板｜GitHub Pages 雲端公開版
==========================================
這是沿用 V3 分類規則的 GitHub Actions + GitHub Pages 部署專案，不是已公開的網址。

建立固定網址：
1. 註冊/登入 https://github.com ，新增 PUBLIC repository，建議名稱 stock-water-100。
2. 不要勾選自動建立 README；將此專案「資料夾內全部檔案」上傳到 main 分支。
   注意：必須包含隱藏資料夾 .github/workflows/update-and-publish.yml；
   Windows 檔案總管可在『檢視』打開隱藏的項目。
   最可靠方法是用 GitHub Desktop，或 git 指令上傳（詳見 README.md）。
3. Repository → Settings → Pages → Build and deployment → Source 選 GitHub Actions。
4. Repository → Settings → Actions → General → Workflow permissions → Read and write permissions。
5. Repository → Actions → Refresh Taiwan Top 100 Water Dashboard → Run workflow → main → Run workflow。
6. 成功後去 Settings → Pages 查網址，通常為 https://使用者名稱.github.io/stock-water-100/。

更新：週一到週五台灣時間15:10於 GitHub 雲端計算和發布。亦可在 Actions 手動 Run workflow。
注意：瀏覽器打開網址本身不會即時執行 Python，而是顯示最近成功發布的收盤分析。
發布頁面公開；請勿把任何帳號密碼或私有資料上傳到公開 repository。
資料源：TAIFEX 名單、Yahoo Finance 歷史價格（依其服務/使用條款自行確認公開用途）。
若下載失敗，Pages 保留上次成功發布的版本；依頁面行情日期檢查新鮮度。
GitHub 公開 repo 長達60天無活動時，排程可能被自動停用，需重新啟用。
