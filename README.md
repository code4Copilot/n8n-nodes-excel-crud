# n8n-nodes-excel-crud

[![npm version](https://badge.fury.io/js/%40your-name%2Fn8n-nodes-excel-crud.svg)](https://www.npmjs.com/package/@your-name/n8n-nodes-excel-crud)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Excel CRUD 社群節點,支援**直接檔案路徑存取**與**完整工作表管理**功能。

## ✨ 主要特色

- 🎯 **雙模式輸入**: 支援檔案路徑 & Binary Data 兩種輸入方式
- 📊 **完整 CRUD**: 對資料列進行新增、讀取、更新、刪除操作
- 📑 **工作表管理**: 建立、刪除、重新命名、複製工作表
- 💾 **自動儲存**: 檔案路徑模式支援自動儲存變更
- 🔍 **進階搜尋**: 支援精確/包含/開頭符合等多種搜尋方式
- 🎯 **智慧列號回傳**: Find Rows 可選擇回傳列號陣列，便於批次操作
- 📋 **動態欄位選單**: File Path 模式下自動載入工作表欄位清單
- 🚀 **高效能**: 使用 ExcelJS 處理大型 Excel 檔案

## 📦 安裝

### 方法 1: 在 n8n 界面中安裝

1. 開啟 n8n 設定
2. 選擇 "Community Nodes"
3. 輸入 `n8n-nodes-excel-crud`
4. 點擊 "Install"

### 方法 2: 使用 npm 安裝

如果您是自行架設 n8n，可以使用 npm 直接安裝：

```bash
# 切換到 n8n 節點目錄
cd ~/.n8n/nodes

# 安裝社群節點
npm install n8n-nodes-excel-crud

# 重新啟動 n8n
```

或者在 Docker 環境中：

```bash
# 在 docker-compose.yml 中設定
environment:
  - N8N_COMMUNITY_PACKAGES_ENABLED=true

# 進入容器安裝
docker exec -it n8n npm install n8n-nodes-excel-crud

# 重新啟動容器
docker restart n8n
```

### 環境設定

確保啟用社群節點功能：

```bash
export N8N_COMMUNITY_PACKAGES_ENABLED=true
```

## 📚 功能說明

### 🔀 輸入模式

#### 1. File Path Mode (檔案路徑模式) ⭐ 推薦

直接指定 Excel 檔案路徑,節點自動讀取和儲存。

**優點:**
- ✅ 無需處理 Binary Data
- ✅ 支援自動儲存變更
- ✅ 適合處理伺服器上的檔案
- ✅ 工作流程更簡潔

**範例:**
```
檔案路徑: /data/sales.xlsx
自動儲存: 開啟
```

#### 2. Binary Data Mode (二進位資料模式)

從前一個節點的 Binary Property 讀取 Excel 檔案。

**適用情境:**
- 從 HTTP Request 下載的檔案
- 從 Email 附件讀取
- 需要在記憶體中處理檔案

---

## 📊 資料列操作 (Row Operations)

### 1. **Append Row** - 附加資料列

附加新資料列到工作表末端。

**參數:**
- `Input Mode`: 選擇 File Path 或 Binary Data
- `File Path` / `Binary Property`: 檔案來源
- `Sheet Name`: 工作表名稱
- `Row Data`: JSON 格式的列資料
- `Auto Save`: 是否自動儲存 (僅 File Path 模式)

**範例 (File Path 模式):**
```json
{
  "filePath": "/data/employees.xlsx",
  "sheetName": "Employees",
  "rowData": {
    "Name": "Alice Wang",
    "Department": "Engineering",
    "Salary": 80000,
    "StartDate": "2024-01-15"
  },
  "autoSave": true
}
```

**輸出:**
```json
{
  "success": true,
  "operation": "appendRow",
  "rowNumber": 25,
  "data": {...},
  "filePath": "/data/employees.xlsx",
  "saved": true
}
```

### 2. **Read Rows** - 讀取資料列

讀取指定範圍的資料列。

**參數:**
- `Start Row`: 起始列號 (預設: 2)
- `End Row`: 結束列號 (0 = 全部)

**範例:**
```json
{
  "filePath": "/data/inventory.xlsx",
  "sheetName": "Products",
  "startRow": 2,
  "endRow": 0  // 讀取全部
}
```

**輸出:**
```json
[
  {
    "ProductID": "P001",
    "Name": "Laptop",
    "Stock": 45,
    "Price": 25000
  },
  {
    "ProductID": "P002",
    "Name": "Mouse",
    "Stock": 120,
    "Price": 500
  }
  // ... 更多資料
]
```

### 3. **Update Row** - 更新資料列

更新指定列的資料。

**範例:**
```json
{
  "filePath": "/data/customers.xlsx",
  "sheetName": "Customers",
  "rowNumber": 5,
  "updatedData": {
    "Email": "newemail@example.com",
    "Phone": "0912-345-678",
    "Status": "Active"
  },
  "autoSave": true
}
```

### 4. **Delete Row** - 刪除資料列

刪除指定的資料列。

**範例:**
```json
{
  "filePath": "/data/temp_data.xlsx",
  "sheetName": "TempRecords",
  "rowNumber": 10,
  "autoSave": true
}
```

### 5. **Find Rows** - 搜尋資料列

依條件搜尋符合的資料列，支援回傳完整資料或僅回傳列號。

**參數:**
- `Search Column`: 搜尋欄位
  - **File Path 模式**: 從下拉選單選擇欄位名稱 (自動載入)
  - **Binary Data 模式**: 手動輸入欄位名稱或字母 (如 "Name" 或 "A")
- `Search Value`: 搜尋值
- `Match Type`: 比對方式
  - **Exact Match**: 完全符合
  - **Contains**: 包含文字
  - **Starts With**: 開頭符合
- `Return Row Numbers`: 是否僅回傳列號陣列
  - **false** (預設): 回傳完整資料列
  - **true**: 僅回傳符合的列號陣列，適合用於後續的更新、插入、刪除操作

**範例 1: 精確搜尋 - 回傳完整資料**
```json
{
  "filePath": "/data/orders.xlsx",
  "sheetName": "Orders",
  "searchColumn": "Status",  // File Path 模式下從下拉選單選擇
  "searchValue": "Pending",
  "matchType": "exact",
  "returnRowNumbers": false
}
```

**輸出 (完整資料):**
```json
[
  {
    "_rowNumber": 3,
    "OrderID": "ORD001",
    "CustomerName": "John Wang",
    "Status": "Pending",
    "Amount": 15000
  },
  {
    "_rowNumber": 7,
    "OrderID": "ORD005",
    "CustomerName": "Mary Chen",
    "Status": "Pending",
    "Amount": 28000
  }
  // ... 更多符合的資料
]
```

**範例 2: 模糊搜尋 - 僅回傳行號**
```json
{
  "filePath": "/data/customers.xlsx",
  "sheetName": "Customers",
  "searchColumn": "CustomerName",
  "searchValue": "Wang",
  "matchType": "contains",
  "returnRowNumbers": true
}
```

**輸出 (行號陣列):**
```json
{
  "rowNumbers": [3, 7, 12, 18, 25]
}
```

**使用情境 - 批次更新範例:**
```json
// 步驟 1: 搜尋需要更新的資料列
{
  "operation": "findRows",
  "searchColumn": "Status",
  "searchValue": "Pending",
  "returnRowNumbers": true
}
// 輸出: { "rowNumbers": [3, 5, 8] }

// 步驟 2: 使用 Loop Over Items 批次更新
{
  "operation": "updateRow",
  "rowNumber": "{{$json.rowNumbers[0]}}",  // 循環使用每個行號
  "updatedData": {
    "Status": "Processing",
    "UpdateDate": "{{$now}}"
  }
}
```

**使用情境 - 批次刪除範例:**
```json
// 步驟 1: 找出要刪除的資料列
{
  "operation": "findRows",
  "searchColumn": "ExpiryDate",
  "searchValue": "2024-01-01",
  "matchType": "startsWith",
  "returnRowNumbers": true
}
// 輸出: { "rowNumbers": [10, 15, 22] }

// 步驟 2: 從後往前刪除 (避免行號偏移)
// 使用 Function Node 反轉陣列: $json.rowNumbers.reverse()
{
  "operation": "deleteRow",
  "rowNumber": "{{$json.rowNumbers[$index]}}"
}
```
```

### 6. **Insert Row** - 插入資料列

在指定位置插入新資料列。

**範例:**
```json
{
  "filePath": "/data/schedule.xlsx",
  "sheetName": "Schedule",
  "rowNumber": 5,
  "rowData": {
    "Date": "2024-12-25",
    "Event": "Christmas Party",
    "Location": "Office",
    "Attendees": 50
  },
  "autoSave": true
}
```

---

## 📑 工作表操作 (Worksheet Operations)

### 1. **Create Worksheet** - 建立工作表

建立新的工作表,可選擇性加入初始資料。

**參數:**
- `New Sheet Name`: 新工作表名稱
- `Initial Data`: 初始資料 (JSON 陣列格式)

**範例:**
```json
{
  "filePath": "/data/report.xlsx",
  "newSheetName": "Q4_2024",
  "initialData": [
    {"Month": "Oct", "Sales": 100000, "Profit": 20000},
    {"Month": "Nov", "Sales": 120000, "Profit": 25000},
    {"Month": "Dec", "Sales": 150000, "Profit": 30000}
  ],
  "autoSave": true
}
```

**輸出:**
```json
{
  "success": true,
  "operation": "createWorksheet",
  "sheetName": "Q4_2024",
  "rowCount": 4,
  "columnCount": 3,
  "filePath": "/data/report.xlsx",
  "saved": true
}
```

### 2. **Delete Worksheet** - 刪除工作表

刪除指定的工作表。

**範例:**
```json
{
  "filePath": "/data/archive.xlsx",
  "worksheetName": "Old_Data",
  "autoSave": true
}
```

### 3. **Rename Worksheet** - 重新命名工作表

將工作表重新命名。

**範例:**
```json
{
  "filePath": "/data/reports.xlsx",
  "worksheetName": "Sheet1",
  "newSheetName": "Sales_2024",
  "autoSave": true
}
```

### 4. **List Worksheets** - 列出所有工作表

取得工作簿中所有工作表的清單。

**參數:**
- `Include Hidden`: 是否包含隱藏的工作表

**範例:**
```json
{
  "filePath": "/data/master.xlsx",
  "includeHidden": false
}
```

**輸出:**
```json
{
  "operation": "listWorksheets",
  "totalSheets": 3,
  "worksheets": [
    {
      "id": 1,
      "name": "Sales",
      "rowCount": 100,
      "columnCount": 8,
      "state": "visible"
    },
    {
      "id": 2,
      "name": "Inventory",
      "rowCount": 250,
      "columnCount": 10,
      "state": "visible"
    },
    {
      "id": 3,
      "name": "Archive",
      "rowCount": 500,
      "columnCount": 12,
      "state": "visible"
    }
  ]
}
```

### 5. **Copy Worksheet** - 複製工作表

複製現有工作表到新工作表。

**範例:**
```json
{
  "filePath": "/data/templates.xlsx",
  "worksheetName": "Template_2024",
  "newSheetName": "Template_2025",
  "autoSave": true
}
```

**輸出:**
```json
{
  "success": true,
  "operation": "copyWorksheet",
  "sourceName": "Template_2024",
  "newName": "Template_2025",
  "rowCount": 50,
  "filePath": "/data/templates.xlsx",
  "saved": true
}
```

### 6. **Get Worksheet Info** - 取得工作表資訊

取得工作表的詳細資訊,包括欄位設定。

**範例:**
```json
{
  "filePath": "/data/database.xlsx",
  "worksheetName": "Users"
}
```

**輸出:**
```json
{
  "operation": "getWorksheetInfo",
  "sheetName": "Users",
  "rowCount": 150,
  "columnCount": 6,
  "actualRowCount": 151,
  "actualColumnCount": 6,
  "state": "visible",
  "columns": [
    {
      "index": 1,
      "letter": "A",
      "header": "UserID",
      "width": 15
    },
    {
      "index": 2,
      "letter": "B",
      "header": "Name",
      "width": 25
    },
    {
      "index": 3,
      "letter": "C",
      "header": "Email",
      "width": 30
    }
    // ... 更多欄位
  ]
}
```

---

## 💡 實用工作流程範例

### 範例 1: 每日資料匯入

```
工作流程: 每日銷售資料匯入

1. Schedule Trigger (每天 09:00)
2. HTTP Request (下載 CSV 報表)
3. Convert to Excel
4. Excel CRUD - Append Row
   - Mode: File Path
   - File Path: /data/sales_master.xlsx
   - Sheet Name: Daily_Sales
   - Auto Save: true
5. Slack (發送完成通知)
```

### 範例 2: 資料清理與整理 - 使用行號批次刪除 🆕

```
工作流程: 清理過期訂單 (改進版)

1. Schedule Trigger (每週日)
2. Excel CRUD - Find Rows
   - File Path: /data/orders.xlsx
   - Sheet Name: Orders
   - Search Column: Status (從下拉選單選擇) 🆕
   - Search Value: "Expired"
   - Match Type: exact
   - Return Row Numbers: true 🆕
   輸出: { "rowNumbers": [25, 42, 68, 103] }

3. Function Node (反轉陣列，避免行號偏移)
   ```javascript
   return {
     json: {
       rowNumbers: $json.rowNumbers.reverse()
     }
   };
   ```
   輸出: { "rowNumbers": [103, 68, 42, 25] }

4. Split Out (展開行號陣列)
   - Field to Split Out: rowNumbers

5. Excel CRUD - Delete Row (循環刪除)
   - Row Number: {{$json}}
   - Auto Save: true

6. Email (寄送報告)
   - 主旨: 已清理 {{$node["Split Out"].itemCount}} 筆過期訂單
```

### 範例 3: 多工作表報表生成

```
工作流程: 月度報表生成

1. Manual Trigger
2. Excel CRUD - Create Worksheet
   - File Path: /data/monthly_report.xlsx
   - New Sheet Name: December_2024
   - Initial Data: (從資料庫查詢)
3. Excel CRUD - Copy Worksheet
   - Source: Template
   - New Name: December_2024_Formatted
4. Excel CRUD - Update Row (格式化資料)
5. Email (寄送報表)
```

### 範例 4: 庫存管理自動化 - 使用行號批次更新 🆕

```
工作流程: 庫存低於閾值警示 (使用新功能)

1. Schedule Trigger (每小時)
2. Excel CRUD - Find Rows
   - File Path: /data/inventory.xlsx
   - Sheet Name: Products
   - Search Column: Status (從下拉選單選擇)
   - Search Value: "Low Stock"
   - Match Type: exact
   - Return Row Numbers: true  🆕
   輸出: { "rowNumbers": [5, 12, 18, 23] }

3. Split Out (將 rowNumbers 陣列展開)
   - Field to Split Out: rowNumbers

4. Excel CRUD - Update Row (循環處理每個行號)
   - Row Number: {{$json}}
   - Updated Data: {"Status": "Reordered", "AlertDate": "{{$now}}"}
   - Auto Save: true

5. Excel CRUD - Read Rows (讀取已更新的資料)
   - 用於取得完整資料供 Email 使用

6. Email (寄送警示給採購部門)
   - 包含已觸發補貨的產品清單
```

### 範例 5: 客戶資料搜尋 API - 支援動態欄位選擇 🆕

```
工作流程: 客戶資料查詢 API (改進版)

1. Webhook (POST /search-customer)
   Request Body:
   {
     "field": "Email",      // 欄位名稱
     "value": "wang",       // 搜尋值
     "matchType": "contains",
     "returnRowNumbers": false
   }

2. Excel CRUD - Find Rows
   - File Path: /data/customers.xlsx
   - Sheet Name: Customers
   - Input Mode: filePath (支援下拉選單選擇欄位) 🆕
   - Search Column: {{$json.body.field}}
   - Search Value: {{$json.body.value}}
   - Match Type: {{$json.body.matchType}}
   - Return Row Numbers: {{$json.body.returnRowNumbers}}

3. Function Node (格式化結果)
   ```javascript
   // 根據 returnRowNumbers 參數決定回傳格式
   if ($input.item.json.rowNumbers) {
     // 回傳行號陣列
     return {
       success: true,
       count: $input.item.json.rowNumbers.length,
       rowNumbers: $input.item.json.rowNumbers
     };
   } else {
     // 回傳完整資料
     return {
       success: true,
       count: $input.all().length,
       data: $input.all()
     };
   }
   ```

4. Respond to Webhook (返回 JSON)
```

### 範例 6: 工作表自動備份

```
工作流程: 每月工作表備份

1. Schedule Trigger (每月 1 日)
2. Excel CRUD - List Worksheets
   - File Path: /data/production.xlsx
3. Loop Over Items
   └─ Excel CRUD - Copy Worksheet
      - Source: {{$json.name}}
      - New Name: {{$json.name}}_backup_{{$now.format('YYYYMM')}}
4. Excel CRUD - Get Worksheet Info (驗證)
5. Slack (發送備份完成通知)
```

---

## 🔧 完整配置範例

### 配置 1: File Path 模式 + 自動儲存

```json
{
  "resource": "row",
  "inputMode": "filePath",
  "filePath": "/data/employees.xlsx",
  "sheetName": "Employees",
  "operation": "appendRow",
  "rowData": {
    "EmployeeID": "EMP123",
    "Name": "Alice Wang",
    "Department": "Engineering",
    "Salary": 80000,
    "JoinDate": "2024-12-01"
  },
  "autoSave": true
}
```

### 配置 2: Binary Data 模式

```json
{
  "resource": "row",
  "inputMode": "binaryData",
  "binaryPropertyName": "data",
  "sheetName": "Orders",
  "operation": "readRows",
  "startRow": 2,
  "endRow": 0
}
```

### 配置 3: 工作表管理

```json
{
  "resource": "worksheet",
  "inputMode": "filePath",
  "filePath": "/data/master.xlsx",
  "worksheetOperation": "createWorksheet",
  "newSheetName": "Q4_2024",
  "initialData": [
    {"Month": "Oct", "Revenue": 100000},
    {"Month": "Nov", "Revenue": 120000},
    {"Month": "Dec", "Revenue": 150000}
  ],
  "autoSave": true
}
```

---

## 🎯 最佳實踐

### 1. 檔案路徑管理

```bash
# 建議的目錄結構
/data/
  ├── excel/
  │   ├── master/        # 主要資料檔案
  │   ├── temp/          # 臨時檔案
  │   ├── archive/       # 封存檔案
  │   └── templates/     # 範本檔案
```

### 2. 錯誤處理

```
建議設定:
- Continue on Fail: 開啟
- Error Workflow: 設定錯誤通知流程
```

### 3. 效能優化

- ✅ 大檔案使用 File Path 模式
- ✅ 批次操作時分批處理
- ✅ 讀取時指定範圍 (避免讀取全部)
- ✅ 適時使用 Auto Save 減少記憶體使用

### 4. 安全建議

```javascript
// 驗證檔案路徑
const allowedPaths = ['/data/excel/', '/home/n8n/data/'];
const filePath = $node["Excel CRUD"].json.filePath;

if (!allowedPaths.some(path => filePath.startsWith(path))) {
  throw new Error('Invalid file path');
}
```

---

## 🛠️ 相容性

- **n8n**: >= 0.175.0
- **Node.js**: >= 18.10
- **Excel 格式**: .xlsx, .xls, .csv

## 📄 授權

MIT License

## 🙏 支援

需要幫助? 
- [GitHub Issues](https://github.com/your-username/n8n-nodes-excel-crud/issues)
- [n8n Community](https://community.n8n.io/)

---

⭐ 如果這個節點對你有幫助,請給我們一個 Star!