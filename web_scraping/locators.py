from selenium.webdriver.common.by import By

class Locators:
    def __init__(self, hosp):
        if hosp == 'fntn':
            self.table = (By.XPATH, "//div[contains(@id, 'content')]//table")
            self.next_button = (By.XPATH, "//div[contains(@class, 'next')]//a[contains(text(), 'Nasled')]")
        elif hosp == 'donsp':
            self.table = (By.XPATH, "//div[contains(@class, 'responsive-table')]//table")
            self.next_button = (By.XPATH, "//div[contains(@class, 'responsive-table')]//table[contains(@class, 'container')]//tr[contains(@class, 'foot')]//a[contains(text(), '»')]")
