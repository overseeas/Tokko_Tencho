from robocorp import windows

def login_tokko(path, id, password):
    desktop = windows.desktop()
    desktop.windows_run(path)
    login = vault.get_secret("tokko")
    tokko_login = windows.find_window('subname:"特攻店長 - ログイン"')
    tokko_login.find("id:txtUserID").set_value(id)
    tokko_login.find("id:txtPassword").set_value(password, validator = None)
    tokko_login.find("id:btnLogin").click()

def select_menu(menu):
    tokko_main = windows.find_window('subname:"特攻店長 - メインメニュー"')
    menu_dict = {
        "受注管理":"id:pbOrder",
        "受注一覧":"id:pbOrderList",
        "新規受注":"id:pbNewOrder",
        "顧客管理":"id:pbCustomer",
        "入金処理":"id:pbMoneyReceived",
        "メールテンプレ管理":"id:pbMailTemplate",
        "店舗別メールテンプレ管理":"id:pbShopMailTemplate",
        "商品管理":"id:pbItem",
        "商品一覧":"id:pbItemList",
        "項目一括変更":"id:pbBatchTableUpdater",
        "商品テンプレ管理":"id:pbItemTemplate",
        "項目別テンプレ設定":"id:pbColumnTemplateSetting",
        "在庫管理":"id:pbStock",
        "在庫一覧":"id:pbStockList",
        "在庫配分":"id:pbStockDistribution",
        "倉庫在庫登録":"id:pbWarehouseStock",
        "外部在庫関連付け":"id:pbOutsideWarehouseStock"
    }
    tokko_main.find(menu_dict[menu]).click()

def set_search_config():
    tokko_order = windows.find_window('subname:"特攻店長 - 受注一覧"')
    #initialize status
    tokko_order.select("", locator="id:cmbStatus")

    #詳細検索
    assert tokko_order.find("name:詳細検索▼").click()
    detailed_options = tokko_order.find_many("id:cmbSearchType", search_strategy="all", search_depth=6)
    detailed_options = detailed_options[:-1]

def option_text(index, name, value, contain = "を含む"):
    tokko_order = windows.find_window('subname:"特攻店長 - 受注一覧"')
    detailed_options = tokko_order.find_many("id:cmbSearchType", search_strategy="all", search_depth=6)
    detailed_options = detailed_options[:-1]
    detailed_options[index].select(name)
    option_panel = detailed_options[index].get_parent()
    option_panel.set_value(value, locator="id:txtValue")
    option_panel.select(contain, locator="id:cmbSearchCondition")

def option_list(index, name, value, contain = "を含む"):
    tokko_order = windows.find_window('subname:"特攻店長 - 受注一覧"')
    detailed_options = tokko_order.find_many("id:cmbSearchType", search_strategy="all", search_depth=6)
    detailed_options = detailed_options[:-1]
    detailed_options[index].select(name)
    option_panel = detailed_options[index].get_parent()
    option_panel.select(value, locator="id:cmbValue")
    option_panel.select(contain, locator="id:cmbSearchCondition")