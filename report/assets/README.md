# 圖片素材放置說明

把截圖以下列檔名放進這個資料夾，然後重跑：

```bash
node report/build_pptx.js
```

腳本會自動抓取。**檔案不存在時不會報錯**，該位置會顯示灰底佔位方塊，
上面寫著該放什麼內容，所以可以先產出簡報、之後再補圖。

支援副檔名：`.png`、`.jpg`、`.jpeg`（依此順序尋找，找到第一個就用）。

## 檔名清單

### 神秘客系統畫面（最重要）

| 檔名 | 出現頁 | 建議內容 |
|---|---|---|
| `mystery-01-login.png` | P20 | 登入頁或系統首頁 |
| `mystery-02-permission.png` | P21 | 能看出權限分層的畫面（如角色管理、不同角色的選單） |
| `mystery-03-question.png` | P22 | 設計新問題的介面 |
| `mystery-04-assign.png` | P22 | 抽選問題並分配至專案組的介面 |
| `mystery-05-scoring.png` | P22 | 分數評鑑介面 |
| `mystery-06-report.png` | P23 | 報表匯出畫面或匯出結果 |
| `mystery-07-review.png` | P23 | 提交檢討報告的介面 |
| `mystery-08-appeal.png` | P23 | 申覆系統介面 |
| `mystery-09-account.png` | P24 | 帳號／身分組管理與移轉介面 |
| `mystery-10-notify.png` | P24 | 通知系統（全服／個人） |

### 舊流程截圖（用來做前後對比，越亂越有說服力）

| 檔名 | 出現頁 | 建議內容 |
|---|---|---|
| `legacy-excel-01.png` | P19 | 舊的 Excel 表單，最好能看出欄位很多很雜 |
| `legacy-excel-02.png` | P19 | 另一份 Excel，用來呈現「同一件事散在很多份」 |
| `legacy-form-01.png` | P19 | 每月要重設的 Google 表單 |
| `legacy-vs-new.png` | P25 | 若有舊新並排的圖可放這裡；沒有則保持空白，該頁改用表格呈現 |

### PRD 機器人

| 檔名 | 出現頁 | 建議內容 |
|---|---|---|
| `prd-02-dialogue.png` | P27 | **最有說服力的一張**：反問過程的對話畫面，最好能看出它一路追問的樣子 |
| `prd-03-output.png` | P28 | 產出的 PRD 初稿或技術文件 |

### 課程教材

| 檔名 | 出現頁 | 建議內容 |
|---|---|---|
| `course-01-slide.png` | P12 | 「AI 時代的商業企劃思維與簡報製作工具」教材頁 |
| `course-02-slide.png` | P13 | 「AI 圖片生成課」教材頁 |
| `course-03-slide.png` | P14 | 「AI 輔助 UI/UX 與原型設計」教材頁 |
| `course-04-slide.png` | P14 | 「AI 輔助 vibe-coding 程式開發」教材頁 |
| `course-showcase-01.png` | P15 | 教材或課堂實況，展示用 |
| `course-showcase-02.png` | P15 | 同上 |
| `course-showcase-03.png` | P15 | 同上 |

## 截圖小提醒

- **去識別化**：截圖裡若有同事姓名、電話、客戶資料，記得先遮蔽再放進來。
- **解析度**：建議寬度至少 1200px，投影出來才不會糊。
- **裁切**：只截需要的區塊，不用整個瀏覽器視窗（網址列、書籤列可以裁掉）。
