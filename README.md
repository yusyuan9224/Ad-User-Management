# Windows AD User Search and Management Tool

This tool is designed to assist IT personnel in companies and enterprises to search for users in Windows AD, manually insert `employeeID`, and export specific user ranges to Excel. This reduces the manual management workload of IT personnel and improves work efficiency.

## Features

1. Search for users in Windows AD.
2. Manually insert `employeeID`.
3. Export specific user ranges to Excel.

## Requirements

- Python 3.x
- `ldap3` module
- `tkinter` module
- `openpyxl` module

## Installation

1. Clone this repository:
    ```sh
    git clone https://github.com/yourusername/ad-user-management.git
    ```
2. Install the required Python packages:
    ```sh
    pip install ldap3 openpyxl
    ```

## Configuration

Before using this tool, configure the `settings.ini` file. Here is an example of `settings.ini`:

```ini
[LDAP]
server = ldaps://192.168.xxx.xxx
user = xxxxxx\helpdesk
password = xxxxxx
search_base = DC=xxx,DC=xxx
```

## Usage

Run the main program to open the GUI and start operating:

```sh
python main.py
```

## Feature Introduction

1. Enter a name or email address in the input box, select the email domain, and click the "Search" button to start searching.
2. The search results will be displayed in a new window where you can view detailed information.
3. Click the "Add EmployeeID" button to manually insert `employeeID`.
4. Select a specific domain and click the "Export" button to export the search results to an Excel file.

## Contributing

Contributions are welcome. Please fork this repository, create your branch, make your changes, and submit a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

**Disclaimer:** This tool is for academic research and internal use only. We are not responsible for any loss or damage caused by using this tool.

# Windows AD 用戶搜尋和管理工具

此工具旨在幫助公司和企業的 IT 人員搜尋 Windows AD 上的用戶，並允許手動插入 `employeeID`，以及匯出特定範圍人員至 Excel。這樣可以減少 IT 人員手動管理的工作量，提高工作效率。

## 功能

1. 搜尋 Windows AD 上的用戶。
2. 手動插入 `employeeID`。
3. 匯出特定範圍人員至 Excel。

## 需求

- Python 3.x
- `ldap3` 模組
- `tkinter` 模組
- `openpyxl` 模組

## 安裝

1. 克隆此倉庫：
    ```sh
    git clone https://github.com/yourusername/ad-user-management.git
    ```
2. 安裝所需的 Python 套件：
    ```sh
    pip install ldap3 openpyxl
    ```

## 配置

在使用此工具之前，請先配置 `settings.ini` 文件。以下是 `settings.ini` 文件範例：

```ini
[LDAP]
server = ldaps://192.168.xxx.xxx
user = xxxxxx\helpdesk
password = xxxxxx
search_base = DC=xxx,DC=xxx
```

## 使用方法

運行主程序，打開 GUI 進行操作：

```sh
python main.py
```

## 功能介紹

1. 在輸入框中輸入姓名或電子郵件地址，選擇信箱網域，然後點擊 "搜尋" 按鈕開始搜尋。
2. 搜尋結果會顯示在新窗口中，可以查看詳細信息。
3. 點擊 "新增 EmployeeID" 按鈕可以手動插入 `employeeID`。
4. 選擇特定的網域，點擊 "匯出" 按鈕將搜尋結果匯出到 Excel 文件。

## 貢獻

歡迎對此項目進行貢獻。請先 fork 此倉庫，創建您的分支，進行修改，然後提交 pull request。

## 許可

此項目採用 MIT 許可證。詳情請參見 [LICENSE](LICENSE) 文件。

---

**免責聲明：** 此工具僅供學術研究和內部使用，不對因使用此工具導致的任何損失或損害負責。
