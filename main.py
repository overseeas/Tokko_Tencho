from robocorp import windows
import time

def open():
    desktop = windows.desktop()
    desktop.windows_run("C:\Program Files (x86)\TokkoTencho\特攻店長\特攻店長.exe")

def login(id, password):
    tokko_login = windows.find_window('subname:"特攻店長 - ログイン"')
    tokko_login.find("id:txtUserID").set_value(id)
    tokko_login.find("id:txtPassword").set_value(password, validator = None)
    tokko_login.find("id:btnLogin").click()
    assert windows.find_window('subname:"特攻店長 - メインメニュー"')

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

def wait(max_time=60):
    #特攻店長 - 進行状況
    for i in range(max_time):
        try:
            windows.find_window('subname:"特攻店長 - 進行状況"', wait_time=1, timeout=1)
        except:
            break

def close_tokko():
    pass
