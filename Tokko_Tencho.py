from robocorp import windows
import re
import time

class Main:
    def open(self):
        desktop = windows.desktop()
        desktop.windows_run("C:\Program Files (x86)\TokkoTencho\特攻店長\特攻店長.exe")

    def login(self, id, password):
        tokko_login = windows.find_window('subname:"特攻店長 - ログイン"')
        tokko_login.find("id:txtUserID").set_value(id)
        tokko_login.find("id:txtPassword").set_value(password, validator = None)
        tokko_login.find("id:btnLogin").click()
        time.sleep(3)
        windows.find_window('subname:"特攻店長 - メインメニュー"')

    def select_menu(self, menu):
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
        time.sleep(3)

    def wait(self, max_time=60):
        #特攻店長 - 進行状況
        for i in range(max_time):
            try:
                windows.find_window('subname:"特攻店長 - 進行状況"', wait_time=1, timeout=1)
            except:
                break

    def close(self):
        tokko = windows.find_window('subname:"特攻店長"')
        tokko.close_window()


class Order:
    def __init__(self) -> None:
        self.ORDER = windows.find_window('subname:"特攻店長 - 受注一覧"')

    def initialize(self):
        #initialize status
        self.ORDER.find("name:クリア id:btnClearSearchForm").click()
        self.ORDER.select("", locator="id:cmbStatus")
        #詳細検索
        if self.ORDER.find_many("name:詳細検索▼"):
            self.ORDER.find("name:詳細検索▼").click()
        
        self.detailed_options = self.ORDER.find_many("id:cmbSearchType", search_strategy="all", search_depth=6)
        self.detailed_options = self.detailed_options[:-1]
        self.detailed_options.reverse()

    def option_text(self, index: int, option_name: str, input_text: str, contain = "を含む"):
        """
        index should be 0~9, which is the position of 詳細検索
        """
        self.detailed_options[index].select(option_name)
        option_panel = self.detailed_options[index].get_parent()
        option_panel.set_value(input_text, locator="id:txtValue")
        option_panel.select(contain, locator="id:cmbSearchCondition")

    def option_list(self, index: int, option_name: str, input_option: str, contain = "を含む"):
        self.detailed_options[index].select(option_name)
        option_panel = self.detailed_options[index].get_parent()
        option_panel.select(input_option, locator="id:cmbValue")
        option_panel.select(contain, locator="id:cmbSearchCondition")

    def wait(self, max_time=60):
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


    def search(self):
        self.ORDER.find(locator='name:"検索" id:btnSearch').click()
        self.wait()
        return int(re.findall(r'\d+', self.ORDER.find(locator="id:lblOrderCount").name)[0])

    def list_all_click(self):
        self.ORDER.find('name:"左上のヘッダー セル"').click()

    def download_file(self, template, path):
        self.ORDER.find("id:cmbOutputFile").select(template)
        self.ORDER.find("name:出力").click()
        #CSVファイルを出力
        self.ORDER.find("id:FileNameControlHost", timeout=5).set_value(path)
        self.ORDER.find("control:ButtonControl id:1").click()
        self.ORDER.find("control:ButtonControl name:OK").click()
        self.wait()

    def input_order_id(self, id):
        self.ORDER.set_value(id, locator="id:txtOrderID")

    def input_order_number(self, number):
        self.ORDER.set_value(number, locator="id:txtOrderNumber")

    def get_values_from_list(self, column_name) -> list:
        """
        column_name: 結果に表示された項目名を指定してください。
        """
        rows = self.ORDER.find("name:DataGridView").find_many("control:CustomControl")
        result = []
        for row in rows:
            for item in row.find_many("control:DataItemControl"):
                if column_name + " " in item.name:
                    result.append(item.get_value())
        return result

    def send_mails(self, template):
        self.ORDER.find("id:btnOrderConfirmationMailSending").click()
        mail = self.ORDER.find_child_window("id:FrmBatchMailCreator")
        mail.find("id:cmbMailKind").select(template)
        mail.find("id:btnMakeMail").click()
        self.wait()
        mail.find("id:btnSendMail").click()

        self.wait(3000)
        self.ORDER.find("control:ButtonControl class:Button path:1|1|1").click()
        self.wait()
        self.ORDER.find("control:ButtonControl class:Button path:1|1|1").click()

        mail.click("name:閉じる")
        self.wait()

    def open_bulk_change(self):
        self.ORDER.find("id:btnEditPlural").click()
        return self.ORDER.find_child_window("id:FrmOrderBatchUpdater")
        
    def inspect(self):
        self.ORDER.inspect()