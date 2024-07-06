import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk, simpledialog
from ldap3 import Server, Connection, ALL, MODIFY_REPLACE
import configparser
import json
import openpyxl
from openpyxl import Workbook

# 讀取配置文件
config = configparser.ConfigParser()
config.read('settings.ini')

# LDAP 伺服器和連接配置
LDAP_SERVER = config.get('LDAP', 'server')
LDAP_USER = config.get('LDAP', 'user')
LDAP_PASSWORD = config.get('LDAP', 'password')
SEARCH_BASE = config.get('LDAP', 'search_base')

# 函數：獲取所有網域
def get_domains():
    server = Server(LDAP_SERVER, get_info=ALL)
    conn = Connection(server, user=LDAP_USER, password=LDAP_PASSWORD, authentication='NTLM')
    if not conn.bind():
        messagebox.showerror("LDAP 連線失敗", f"無法連接至 LDAP 服務器: {conn.result}")
        return []

    search_filter = "(mail=*)"
    conn.search(SEARCH_BASE, search_filter, attributes=['mail'])
    domains = set()
    for entry in conn.entries:
        for mail in entry.mail:
            domain = mail.split('@')[-1]
            domains.add(domain)
    conn.unbind()
    return list(domains)

# 函數：搜尋 AD 用戶
def search_ad(query, domain):
    print(f"Searching for query: {query} in domain: {domain}")
    server = Server(LDAP_SERVER, get_info=ALL)
    conn = Connection(server, user=LDAP_USER, password=LDAP_PASSWORD, authentication='NTLM')
    if not conn.bind():
        messagebox.showerror("LDAP 連線失敗", f"無法連接至 LDAP 服務器: {conn.result}")
        return []

    search_attributes = ['cn', 'department', 'mail', 'description', 'displayName', 'physicalDeliveryOfficeName', 'employeeID']

    if domain == "All":
        if query:
            search_filter = f"(|(displayName=*{query}*)(mail=*{query}*))"
        else:
            search_filter = "(|(displayName=*)(mail=*))"
    else:
        if query:
            search_filter = f"(|(displayName=*{query}*)(mail=*{query}@{domain}))"
        else:
            search_filter = f"(mail=*@{domain})"
    
    print(f"Using search filter: {search_filter}")

    try:
        conn.search(SEARCH_BASE, search_filter, attributes=search_attributes)
        results = [entry for entry in conn.entries]
    except Exception as e:
        messagebox.showerror("搜尋錯誤", f"搜尋過程中發生錯誤: {str(e)}")
        results = []
    finally:
        conn.unbind()
    
    return results

# 函數：格式化搜尋結果
def format_result(result):
    formatted_result = ""
    for key, value in result.entry_attributes_as_dict.items():
        formatted_result += f"{key}: {value}\n"
    return formatted_result

# 函數：格式化用戶詳細信息
def format_details(entry):
    details = (
        f"name: {entry.cn[0] if 'cn' in entry else 'N/A'}\n"
        f"department: {entry.department[0] if 'department' in entry else 'N/A'}\n"
        f"mail: {entry.mail[0] if 'mail' in entry else 'N/A'}\n"
        f"description: {entry.description[0] if 'description' in entry else 'N/A'}\n"
        f"displayName: {entry.displayName[0] if 'displayName' in entry else 'N/A'}\n"
        f"physicalDeliveryOfficeName: {entry.physicalDeliveryOfficeName[0] if 'physicalDeliveryOfficeName' in entry else 'N/A'}\n"
        f"employeeID: {entry.employeeID[0] if 'employeeID' in entry else 'N/A'}\n"
    )
    return details

# 函數：處理搜尋請求
def on_search():
    global search_results
    query = entry.get().strip()
    domain = domain_var.get()
    if not query:
        messagebox.showwarning("輸入錯誤", "請輸入姓名或電子郵件地址")
        return
    
    search_results = search_ad(query, domain)
    if search_results:
        for result in search_results:
            result_text = format_result(result)
            show_result_window(result_text, result)
    else:
        messagebox.showinfo("搜尋結果", "未找到任何結果")

# 函數：顯示搜尋結果窗口
def show_result_window(result_text, result):
    result_window = tk.Toplevel(root)
    result_window.title("搜尋結果")
    
    text_area = scrolledtext.ScrolledText(result_window, wrap=tk.WORD, width=100, height=10)
    text_area.pack(padx=10, pady=10)
    text_area.insert(tk.END, result_text)
    text_area.configure(state='disabled')

    details_text = format_details(result)
    details_area = scrolledtext.ScrolledText(result_window, wrap=tk.WORD, width=100, height=10)
    details_area.pack(padx=10, pady=10)
    details_area.insert(tk.END, details_text)
    details_area.configure(state='disabled')

    add_button = tk.Button(result_window, text="新增 EmployeeID", command=lambda: confirm_add_employee_id(result))
    add_button.pack(pady=10)

# 函數：確認新增 EmployeeID
def confirm_add_employee_id(entry):
    if hasattr(entry, 'description') and entry.description:
        employee_id = entry.description[0].split(' ')[0]
        response = messagebox.askquestion("確認新增", f"要新增的 EmployeeID: {employee_id}\n\n請確認是否新增？", icon='warning')
        if response == 'yes':
            add_employee_id(entry, employee_id)
        else:
            manual_employee_id = simpledialog.askstring("手動新增", "請輸入 EmployeeID:")
            if manual_employee_id:
                add_employee_id(entry, manual_employee_id)
    else:
        manual_employee_id = simpledialog.askstring("手動新增", "請輸入 EmployeeID:")
        if manual_employee_id:
            add_employee_id(entry, manual_employee_id)

# 函數：新增 EmployeeID
def add_employee_id(entry, employee_id):
    server = Server(LDAP_SERVER, get_info=ALL)
    conn = Connection(server, user=LDAP_USER, password=LDAP_PASSWORD, authentication='NTLM')
    if not conn.bind():
        messagebox.showerror("LDAP 連線失敗", f"無法連接至 LDAP 服務器: {conn.result}")
        return

    dn = entry.entry_dn
    conn.modify(dn, {'employeeID': [(MODIFY_REPLACE, [employee_id])]})
    
    conn.unbind()
    messagebox.showinfo("成功", "已成功新增 EmployeeID")

# 函數：匯出搜尋結果到 Excel
def export_to_excel():
    domain = domain_var.get()
    
    if domain == "All":
        messagebox.showwarning("輸入錯誤", "請選擇特定的網域")
        return
    
    print(f"Exporting results for domain: {domain}")
    results = search_ad("", domain)
    if not results:
        messagebox.showwarning("錯誤", "未找到任何結果")
        return
    
    wb = Workbook()
    ws = wb.active
    ws.append(['name', 'department', 'mail', 'description', 'displayName', 'physicalDeliveryOfficeName', 'employeeID'])
    
    for entry in results:
        row = [
            entry.cn[0] if 'cn' in entry else '',
            entry.department[0] if 'department' in entry else '',
            entry.mail[0] if 'mail' in entry else '',
            entry.description[0] if 'description' in entry else '',
            entry.displayName[0] if 'displayName' in entry else '',
            entry.physicalDeliveryOfficeName[0] if 'physicalDeliveryOfficeName' in entry else '',
            entry.employeeID[0] if 'employeeID' in entry else ''
        ]
        ws.append(row)
    
    file_name = "exported_results.xlsx"
    wb.save(file_name)
    messagebox.showinfo("成功", f"已成功匯出到 {file_name}")

# 創建主窗口
root = tk.Tk()
root.title("AD 用戶搜尋工具")

# 創建並放置控件
tk.Label(root, text="輸入姓名或電子郵件地址:").grid(row=0, column=0, padx=10, pady=10)
entry = tk.Entry(root, width=40)
entry.grid(row=0, column=1, padx=10, pady=10)

search_button = tk.Button(root, text="搜尋", command=on_search)
search_button.grid(row=0, column=3, padx=10, pady=10)

# 獲取AD中的所有網域
domains = get_domains()
domains.insert(0, "All")  # 添加 "All" 選項

# 添加下拉選單
tk.Label(root, text="選擇信箱網域:").grid(row=1, column=0, padx=10, pady=10)
domain_var = tk.StringVar(value="All")
domain_menu = ttk.Combobox(root, textvariable=domain_var, values=domains, state='readonly')
domain_menu.grid(row=1, column=1, padx=10, pady=10)

# 添加匯出按鈕
export_button = tk.Button(root, text="匯出", command=export_to_excel)
export_button.grid(row=2, column=1, padx=10, pady=10)

# 開始主事件循環
root.mainloop()
