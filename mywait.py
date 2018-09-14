import time


class mywait():
    def __init__(self, seleniumobj,waittime = 10):
        self.seleniumobj = seleniumobj
        self.waittime = waittime
        self.cursor = None
    def find_single_xpath(self, xpath):
        for i in range(0,self.waittime):
            try:
                element = self.seleniumobj.find_element_by_xpath(xpath)
                if element is None:
                    time.sleep(1)
                else:
                    break
            except:
                time.sleep(1)
                if i == self.waittime - 1:
                    raise RuntimeError("wait too long, selenium failed")

        return element

    def find_mul_xpath(self, xpath):

        for i in range(0, self.waittime):
            try:
                elements = self.seleniumobj.find_elements_by_xpath(xpath)
                if len(elements) == 0:
                    time.sleep(1)
                else:
                    break
            except:
                time.sleep(1)
                if i == self.waittime - 1:
                    raise RuntimeError("wait too long, selenium failed")

        return elements
    def clear(self):
        for i in range(0, self.waittime):
            try:
                self.seleniumobj.clear()
                break
            except:
                time.sleep(1)
                if i == self.waittime - 1:
                    raise RuntimeError("wait too long, selenium failed")

        return None

    def send_keys(self,content):
        for i in range(0, self.waittime):
            try:
                self.seleniumobj.send_keys(content)
                break
            except:
                time.sleep(1)
                if i == self.waittime - 1:
                    raise RuntimeError("wait too long, selenium failed")

        return None

    def click(self,cursor):
        for i in range(0, self.waittime):
            try:
                self.cursor = cursor
                self.seleniumobj.execute_script("arguments[0].click();",self.cursor)
                break
            except:
                time.sleep(1)
                if i == self.waittime - 1:
                    raise RuntimeError("wait too long, selenium failed")

        return None

    def select_by_value(self,content):
        for i in range(0, self.waittime):
            try:
                self.seleniumobj.select_by_value(content)
                break
            except:
                time.sleep(1)
                if i == self.waittime - 1:
                    raise RuntimeError("wait too long, selenium failed")

        return None


