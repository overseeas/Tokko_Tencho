from robocorp import windows
import re

def initialize():
    tokko_order = windows.find_window('subname:"特攻店長 - 受注一覧"')
    #initialize status
    tokko_order.select("", locator="id:cmbStatus")
    #詳細検索
    if tokko_order.find_many("name:詳細検索▼"):
        tokko_order.find("name:詳細検索▼").click()
    
    #detailed_options = tokko_order.find_many("id:cmbSearchType", search_strategy="all", search_depth=6)
    #detailed_options = detailed_options[:-1]

def option_text(index: int, option_name: str, input_text: str, contain = "を含む"):
    tokko_order = windows.find_window('subname:"特攻店長 - 受注一覧"')
    detailed_options = tokko_order.find_many("id:cmbSearchType", search_strategy="all", search_depth=6)
    detailed_options = detailed_options[:-1]
    detailed_options[index].select(option_name)
    option_panel = detailed_options[index].get_parent()
    option_panel.set_value(input_text, locator="id:txtValue")
    option_panel.select(contain, locator="id:cmbSearchCondition")

def option_list(index: int, option_name: str, input_option: str, contain = "を含む"):
    tokko_order = windows.find_window('subname:"特攻店長 - 受注一覧"')
    detailed_options = tokko_order.find_many("id:cmbSearchType", search_strategy="all", search_depth=6)
    detailed_options = detailed_options[:-1]
    detailed_options[index].select(option_name)
    option_panel = detailed_options[index].get_parent()
    option_panel.select(input_option, locator="id:cmbValue")
    option_panel.select(contain, locator="id:cmbSearchCondition")

def search():
    tokko_order = windows.find_window('subname:"特攻店長 - 受注一覧"')
    #tokko_order.inspect()
    tokko_order.find(locator='name:"検索" id:btnSearch').click()
    return int(re.findall(r'\d+', tokko_order.find(locator="id:lblOrderCount").name)[0])