from robocorp import windows
import re

def initialize():
    ORDER = windows.find_window('subname:"特攻店長 - 受注一覧"')
    #initialize status
    ORDER.select("", locator="id:cmbStatus")
    #詳細検索
    if ORDER.find_many("name:詳細検索▼"):
        ORDER.find("name:詳細検索▼").click()
    
    #detailed_options = ORDER.find_many("id:cmbSearchType", search_strategy="all", search_depth=6)
    #detailed_options = detailed_options[:-1]

def option_text(index: int, option_name: str, input_text: str, contain = "を含む"):
    ORDER = windows.find_window('subname:"特攻店長 - 受注一覧"')
    detailed_options = ORDER.find_many("id:cmbSearchType", search_strategy="all", search_depth=6)
    detailed_options = detailed_options[:-1]
    detailed_options[index].select(option_name)
    option_panel = detailed_options[index].get_parent()
    option_panel.set_value(input_text, locator="id:txtValue")
    option_panel.select(contain, locator="id:cmbSearchCondition")

def option_list(index: int, option_name: str, input_option: str, contain = "を含む"):
    ORDER = windows.find_window('subname:"特攻店長 - 受注一覧"')
    detailed_options = ORDER.find_many("id:cmbSearchType", search_strategy="all", search_depth=6)
    detailed_options = detailed_options[:-1]
    detailed_options[index].select(option_name)
    option_panel = detailed_options[index].get_parent()
    option_panel.select(input_option, locator="id:cmbValue")
    option_panel.select(contain, locator="id:cmbSearchCondition")

def wait(max_time=60):
    #受注検索中
    for i in range(max_time):
        try:
            windows.find_window('subname:"受注検索中"', wait_time=1, timeout=1)
        except:
            break
    for i in range(max_time):
        try:
            windows.find_window('name:"特攻店長 - 進行状況"', wait_time=1, timeout=1)
        except:
            break


def search():
    ORDER = windows.find_window('subname:"特攻店長 - 受注一覧"')
    #ORDER.inspect()
    ORDER.find(locator='name:"検索" id:btnSearch').click()
    wait()
    return int(re.findall(r'\d+', ORDER.find(locator="id:lblOrderCount").name)[0])

def list_all_click():
    ORDER = windows.find_window('subname:"特攻店長 - 受注一覧"')
    ORDER.find('name:"左上のヘッダー セル"').click()

def download_file(template, path):
    ORDER = windows.find_window('subname:"特攻店長 - 受注一覧"')
    ORDER.find("id:cmbOutputFile").select(template)
    ORDER.find("name:出力").click()
    #CSVファイルを出力
    ORDER.find("id:FileNameControlHost", timeout=5).set_value(path)
    ORDER.find("control:ButtonControl id:1").click()
    ORDER.find("control:ButtonControl name:OK").click()
    wait()

def option_id(id):
    ORDER = windows.find_window('subname:"特攻店長 - 受注一覧"')
    ORDER.set_value(id, locator="id:txtOrderID")

def option_number(number):
    ORDER = windows.find_window('subname:"特攻店長 - 受注一覧"')
    ORDER.set_value(number, locator="id:txtOrderNumber")

def get_value_from_list(column_name):
    ORDER = windows.find_window('subname:"特攻店長 - 受注一覧"')
    rows = ORDER.find("name:DataGridView").find_many("control:CustomControl")
    result = []
    for row in rows:
        for item in row.find_many("control:DataItemControl"):
            if column_name + " " in item.name:
                result.append(item.get_value())
    return result

def send_mails(template):
    ORDER = windows.find_window('subname:"特攻店長 - 受注一覧"')
    #ORDER.inspect()
    ORDER.find("id:btnOrderConfirmationMailSending").click()
    ORDER.find("id:cmbMailKind").select(template)
    ORDER.find("id:btnMakeMail").click()
    ORDER.find("id:btnSendMail").click()
    wait()
    ORDER.find("control:ButtonControl class:Button path:1|1|1").click()
    wait()
    ORDER.find("control:ButtonControl class:Button path:1|1|1").click()
    ORDER.find_child_window("id:FrmBatchMailCreator").click("name:閉じる")
    wait()
