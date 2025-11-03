# PDF 摘要擷取工具（PDF Abstract Extractor）

從學術 PDF（中文/英文）擷取「摘要/Abstract」區段，輸出每篇的 TXT 檔，並同時產生一份彙整用的 CSV。

## 快速開始（Quick Start）

```powershell


# 擷取單一 PDF（輸出到 .\abstract_output）
python pdf_abstract.py ./pdf_example/利用基於聯邦學習的條件生成對抗網路.pdf --output-dir ./abstract_output -v

# 擷取整個資料夾（非遞迴）
python pdf_abstract.py .\pdf_example --output-dir .\abstract_output -v

# 遞迴處理子資料夾
python pdf_abstract.py .\pdf_example --recursive --output-dir .\abstract_output -v

# 僅檢視前幾頁/行（除錯用，已整合在同一程式）
python pdf_abstract.py .\pdf_example\標準4_利用RBF-DNN配合醫院病患人流資料探究.pdf --inspect --inspect-pages 1 --inspect-lines 40
```

## 功能
- 偵測摘要標題：`摘要`、`中文摘要`、`Abstract`、`ABSTRACT`
- 停止於常見結尾：`關鍵字/关键词/Keywords`、`References/參考文獻/参考文献`、或第一章標題（例如 `Introduction/引言/緒論`）
- 支援同一行內的 `摘要：本研究…` 形式
- 跨頁摘要擷取，並設有頁數上限保護（預設 3 頁）
- 每篇輸出 TXT：包含四個欄位「題目／年分／作者／摘要」
- 產生彙整 CSV（`論文整理.csv`）：欄位為「題目、年分、作者、摘要」
- 題目清理：移除檔名中的目錄型前綴（如 `基礎1_`、`標準10_`、`查詢3_`）與尾端日期
- 題目擴充（中文）：
	 - 上位字串擴充：在摘要前區域，若發現「包含檔名題目且更長的中文行」，優先採用該行作為更完整題目（允許少量標點，末尾不得為句號/冒號）。
	 - 右側語境延伸：若檔名題目被拆成多行（例如主標＋副標在下一行），會跨越一個換行拼接為完整題目（僅在下一行看起來像標題延續時）。
	 - 題目會自動移除內部換行，統一為單行輸出，避免 CSV 欄位斷行。
- 詳細模式：
	- `-v`：在主控台輸出每個檔案「作者偵測」的一行摘要（來源標記如 `label-3p/label-8p/guess-3p/guess-8p/metadata/cjk-override`）。
	- `-vv`：更詳細的追蹤（例如題目/作者判斷、摘要起訖偵測訊息）。
	- 作者輸出策略：若作者同時含中英文，預設僅保留中文姓名（便於彙整與一致呈現）。

## CLI 選項

- `input`：單一 PDF 檔或資料夾路徑。
- `--output-dir`：輸出資料夾（預設 `abstract_output`）。
- `--recursive`：若 `input` 是資料夾，啟用遞迴掃描子資料夾。
- `-v` / `--verbose`：精簡紀錄；每篇列印一行作者摘要。
- `-vv` / `--very-verbose`：完整追蹤（包含題目/作者決策、摘要起訖標記等）；隱含 `-v`。
- `--inspect`：僅輸出前若干頁/行的純文字以便除錯，不進行抽取與輸出。
- `--inspect-pages` / `--inspect-lines`：搭配 `--inspect` 控制輸出頁數與行數。



## 使用方式

針對單一檔案：

```powershell
python pdf_abstract.py .\pdf_example\基礎1_以純Youbike資料識別Covid-19對臺北市區活動的影響_20220219.pdf --output-dir .\abstract_output -v
```

針對整個資料夾：

```powershell
python pdf_abstract.py .\pdf_example --output-dir .\abstract_output -v
```

包含子資料夾（遞迴）：

```powershell
python pdf_abstract.py .\pdf_example --recursive --output-dir .\abstract_output -v
```

輸出內容：
- 單篇 TXT：儲存在 `<output-dir>/txt/*.txt`（採用 UTF-8 with BOM，方便 Windows 編輯器）
- 彙整 CSV：儲存為 `<output-dir>/論文整理.csv`（UTF-8 with BOM；使用單一 `\n` 換行，避免 Windows 檢視器顯示空白列）

每個 TXT 檔的格式為四段中文標籤（段落間以空行分隔）：

```
題目：<paper title>

年分：<year>

作者：<author>

摘要：<abstract as a single paragraph>
```

### 詳細輸出（-v / -vv）

加上 `-v` 時，會為每個 PDF 印出一行作者偵測結果；`-vv` 會顯示更完整的偵測過程（亦包含摘要起訖偵測）。格式：

```
[AUTHOR] <作者名稱> (<來源>)
```

來源說明：
- `label-3p`/`label-8p`：在前 3/8 頁內，用「作者/研究生/Author…」等標籤直接擷取
- `guess-3p`/`guess-8p`：在前 3/8 頁內，依姓名外觀與上下文（系所/指導等）推測
- `cjk-override`：題目為中文且作者為純英文時，再以中文姓名候選覆寫
- `metadata`：以上都找不到時，回退 PDF metadata 的作者欄

作者欄位：若偵測到同一行同時出現中英文姓名，輸出時僅保留中文姓名（例如「吳昺儒 Bing-Ru Wu」會輸出為「吳昺儒」）。

註：未加 `-v` 時，僅輸出必要訊息（錯誤/警告）；成功案件會安靜完成，適合批次處理。


## 說明與限制
- 這個工具採用啟發式規則；若 PDF 為掃描影像（無可擷取文字），需要先經過 OCR（例如 Tesseract）。本工具不含 OCR。
- 可於 `pdf_abstract.py` 中調整起訖規則與頁數上限（`START_PATTERNS`、`END_PATTERNS`、`MAX_ABSTRACT_PAGES`）。
- 實際抽取時會保留原始換行以利判斷，但在輸出 TXT 的 `摘要：` 欄位會整併為單一段落，便於閱讀與彙整。
- 若 CSV 檔案當下被開啟（鎖住無法覆寫），程式會顯示警告並略過寫入；請關閉該檔案後重新執行，以更新 `論文整理.csv`。

## 疑難排解（Troubleshooting）
- 若在 Windows 安裝失敗，請先更新 `pip` 後重試。PyMuPDF（pymupdf）的 wheel 已於 PyPI 提供。
- 若遇到不尋常的標題格式造成抽取失敗，可將自訂樣式加入 `START_PATTERNS`／`END_PATTERNS`。
