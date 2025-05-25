# PPT 簡體中文轉繁體中文轉換工具

一個簡單易用的 Python 腳本，用於將 PowerPoint (.pptx) 檔案中的簡體中文轉換為繁體中文，同時保留所有原始格式和樣式。

## 專案背景

這個專案是為了解決使用 AI 工具（如 Kimi）生成簡體中文 PPT 後，需要手動轉換為繁體中文的困擾。

### 為什麼需要這個工具？

1. 使用 Kimi 等 AI 工具生成的 PPT 內容預設為簡體中文
2. 手動轉換每個文字方塊中的簡體字非常耗時
3. 需要一個自動化解決方案來快速完成這項重複性工作

### 這個工具如何幫助你？

- 自動批次處理多個 PPTX 檔案
- 保留原始格式和樣式
- 支援文字方塊、表格和備註中的文字轉換
- 可指定輸出目錄，不會覆蓋原始檔案

## 功能特點

- ✅ 保留原始 PPT 的所有格式和樣式
- ✅ 支援文字方塊、表格和備註中的文字轉換
- ✅ 遞迴處理目錄下的所有 PPTX 檔案
- ✅ 可自訂輸入/輸出目錄
- ✅ 進度顯示和錯誤處理

## 安裝

1. 克隆此儲存庫：
   ```bash
   git clone https://github.com/yourusername/ppt-cn2tw-converter.git
   cd ppt-cn2tw-converter
   ```

2. 安裝所需套件：
   ```bash
   pip install -r requirements.txt
   ```

## 使用方法

### 基本用法

```bash
# 轉換指定目錄中的所有 PPTX 檔案
python3 ppt_cn2tw.py ./input

# 指定輸出目錄
python3 ppt_cn2tw.py ./out -o ./output
```

### 參數說明

```
<目錄路徑>   要處理的目錄路徑（必要參數）
-o, --output   指定輸出目錄（可選，預設為原始目錄）
```

### 範例

```bash
# 轉換 input 目錄中的所有 PPTX 檔案，輸出到 output 目錄
python3 ppt_cn2tw.py ./input -o ./output

# 轉換當前目錄下的所有 PPTX 檔案
python3 ppt_cn2tw.py .
```

## 專案結構

```
.
├── README.md          # 本文件
├── requirements.txt   # 依賴套件
├── src/               # 原始碼目錄
│   └── ppt_cn2tw.py  # 主程式
├── input/            # 預設輸入目錄
└── output/           # 預設輸出目錄
```

## 範例

### 轉換單個檔案

```bash
python src/ppt_cn2tw.py -i ./input/presentation.pptx -o ./output/presentation_tw.pptx
```

### 轉換目錄中的所有檔案

```bash
# 建立輸出目錄
mkdir -p output

# 轉換 input 目錄中的所有 PPTX 檔案
python src/ppt_cn2tw.py -d ./input -o ./output
```

### 遞迴轉換目錄中的檔案

```bash
python src/ppt_cn2tw.py -d ./input -o ./output -r
```

## 授權

[MIT License](LICENSE)

## 貢獻

歡迎提交 Issues 和 Pull Requests！

## 已知限制

- 僅支援 .pptx 格式（PowerPoint 2007 及以上版本）
- 某些特殊字元可能無法正確轉換
- 不支援密碼保護的 PPTX 檔案

## 依賴

- Python 3.6+
- python-pptx
- opencc-python-reimplemented

## 作者

[您的名字] - [您的網站或電子郵件]
