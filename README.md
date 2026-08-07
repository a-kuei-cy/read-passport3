# 興嘉國小電子閱讀護照－前台、學生任務與管理後台

本專案採用 GitHub Pages + Google Apps Script + Google 試算表，包含三個入口：

- `index.html`：公開閱讀成果、排行榜、統計圖表與公開心得。
- `student-tasks.html`：學生查看年級／班級指定書籍、撰寫及重新繳交心得。
- `management.html`：管理者與班級教師登入、編修、審核、指定任務及匯出 Excel。

## 權限與流程

### 系統管理者

- 查看及編修全校閱讀紀錄、學生、書籍與閱讀任務。
- 指定「全年級」或「特定班級」閱讀書籍及心得要求。
- 審核任何班級學生的閱讀心得。
- 匯出閱讀紀錄、學生資料、閱讀任務為 `.xlsx`；Excel 元件無法載入時自動改存 CSV。

### 班級教師

- 只能查看自己班級的學生與閱讀紀錄。
- 審核自己班級學生的閱讀心得。
- 只能建立、編輯及刪除自己班級的指定閱讀任務。
- 可匯出自己班級的閱讀紀錄、學生資料與任務。

### 學生

- 使用學號與班級登入任務頁。
- 同時看見管理者指定給全年級的任務，以及導師指定給本班的任務。
- 填寫閱讀頁數、心得標題與心得全文。
- 被退回後可依教師回饋修改並重新送出。

## 篇數累計規則

每一筆閱讀心得只有在狀態成為 `approved` 時才累計 1 篇。`pending` 與 `rejected` 不計入。學生重新繳交同一個任務會更新原紀錄，不會重複新增篇數。

## 安裝步驟

1. 建立一份 Google 試算表。
2. 開啟「擴充功能 → Apps Script」。
3. 將 `apps-script/Code.gs` 全部貼入 Apps Script。
4. 執行一次 `setupSpreadsheet()`，授權後會建立所需工作表與欄位。
5. 到 `Settings` 工作表修改：
   - `adminKey`：管理者登入金鑰。
   - `teacherKeys`：班級與教師金鑰 JSON，例如：

```json
{"一年一班":"101-key","三年一班":"301-key","六年二班":"602-key"}
```

6. Apps Script 選擇「部署 → 新增部署作業 → 網頁應用程式」。
7. 執行身分選「我」，存取權選可使用此系統的範圍。若 GitHub Pages 公開存取 API，通常需選「任何人」。
8. 複製 `/exec` 網址，貼入 `config.js` 的 `API_URL`。
9. 將整個資料夾上傳到 GitHub repository，啟用 GitHub Pages。
10. 將 `DEMO_MODE` 改為 `false`，公開首頁就會讀取正式試算表資料。

## 建議網址

- 公開成果：`https://帳號.github.io/專案/index.html`
- 學生任務：`https://帳號.github.io/專案/student-tasks.html`
- 管理後台：`https://帳號.github.io/專案/management.html`

## 試算表資料注意事項

- `Students.className`、教師登入班級、任務班級名稱需一致，例如都使用「三年一班」。
- 年級任務的 `scopeType` 為 `grade`，`scopeValue` 可填「3年級」或「三年級」。
- 班級任務的 `scopeType` 為 `class`，`scopeValue` 填完整班級名稱。
- 學生必須先存在 `Students` 工作表且 `active` 不為 `false`，才能登入任務頁。
- 公開心得還需要 `isPublic=TRUE`，才會出現在成果首頁。

## 安全提醒

目前為學校內部輕量版，以管理金鑰進行角色驗證。正式大規模使用時，建議改用 Google Workspace OAuth、校內帳號白名單與後端操作紀錄；不可在前端程式保存管理金鑰。
