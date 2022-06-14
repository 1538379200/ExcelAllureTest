from selenium import webdriver


class BaseControl:
    def __init__(self, type_):
        if type_ == "Chrome":
            self.driver = webdriver.Chrome()
            self.driver.implicitly_wait(10)
            self.driver.maximize_window()
        else:
            raise ValueError("暂时不支持其他浏览器使用！！！")

    def event_open(self, url):
        self.driver.get(url)

    def locate(self, name, value):
        return self.driver.find_element(name, value)

    def event_input(self, name, value, txt):
        self.locate(name, value).send_keys(txt)

    def event_click(self, name, value):
        self.locate(name, value).click()

    # def event_assert(self, name, value, assert_text):
    #     element_text = self.locate(name, value).text
    #     try:
    #         assert assert_text == element_text, "不相等"
    #         return True
    #     except AssertionError:
    #         return False

    def event_assert(self, text):
        try:
            assert "不是你好" == text, "断言错误"
            return True
        except AssertionError:
            return False

    def event_quite(self):
        self.driver.quit()
