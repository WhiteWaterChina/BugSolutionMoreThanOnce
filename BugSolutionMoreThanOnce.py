#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import wx
import time
import os
from threading import Thread
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
import xlsxwriter
import datetime


class BugSolutionMoreThanOnce(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"BUG一次解决率统计", pos=wx.DefaultPosition,
                          size=wx.Size(329, 388), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOWFRAME))

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        self.m_panel1 = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        bSizer5 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText1 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"请输入RDM的用户名", wx.DefaultPosition,
                                           wx.Size(150, 20), 0)
        self.m_staticText1.Wrap(-1)
        self.m_staticText1.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText1.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer5.Add(self.m_staticText1, 0, wx.ALL, 5)

        self.input_username = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          0)
        bSizer5.Add(self.input_username, 1, wx.ALL, 5)

        bSizer3.Add(bSizer5, 0, wx.EXPAND, 5)

        bSizer7 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText2 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"请输入RDM的密码", wx.DefaultPosition, wx.Size(150, 20),
                                           0)
        self.m_staticText2.Wrap(-1)
        self.m_staticText2.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText2.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer7.Add(self.m_staticText2, 0, wx.ALL, 5)

        self.input_password = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          wx.TE_PASSWORD)
        bSizer7.Add(self.input_password, 1, wx.ALL, 5)

        bSizer3.Add(bSizer7, 0, wx.EXPAND, 5)

        bSizer8 = wx.BoxSizer(wx.VERTICAL)

        self.button_login = wx.Button(self.m_panel1, wx.ID_ANY, u"请按这个按钮登录并获取项目信息", wx.DefaultPosition, wx.DefaultSize,
                                      0)
        bSizer8.Add(self.button_login, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer3.Add(bSizer8, 0, wx.EXPAND, 5)

        bSizer9 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText3 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"请在如下选择需要统计BUG一次解决率的项目名称", wx.DefaultPosition,
                                           wx.DefaultSize, 0)
        self.m_staticText3.Wrap(-1)
        self.m_staticText3.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText3.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer9.Add(self.m_staticText3, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        combox_project_listChoices = []
        self.combox_project_list = wx.ComboBox(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                               wx.DefaultSize, combox_project_listChoices, wx.CB_READONLY)
        bSizer9.Add(self.combox_project_list, 0, wx.ALL | wx.EXPAND, 5)

        bSizer3.Add(bSizer9, 0, wx.EXPAND, 5)

        bSizer10 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText4 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"在上面选择项目之后，按GO开始抓取BUG一次解决率", wx.DefaultPosition,
                                           wx.DefaultSize, 0)
        self.m_staticText4.Wrap(-1)
        self.m_staticText4.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText4.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer10.Add(self.m_staticText4, 0, wx.ALL, 5)

        bSizer11 = wx.BoxSizer(wx.HORIZONTAL)

        self.button_go = wx.Button(self.m_panel1, wx.ID_ANY, u"GO", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer11.Add(self.button_go, 0, wx.ALL, 5)

        self.button_close = wx.Button(self.m_panel1, wx.ID_ANY, u"Exit", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer11.Add(self.button_close, 0, wx.ALL, 5)

        bSizer10.Add(bSizer11, 1, wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer3.Add(bSizer10, 0, wx.EXPAND, 5)

        bSizer101 = wx.BoxSizer(wx.VERTICAL)

        self.output_info = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                       wx.TE_READONLY | wx.TE_MULTILINE)
        bSizer101.Add(self.output_info, 1, wx.ALL | wx.EXPAND, 5)

        bSizer3.Add(bSizer101, 1, wx.EXPAND, 5)

        self.m_panel1.SetSizer(bSizer3)
        self.m_panel1.Layout()
        bSizer3.Fit(self.m_panel1)
        bSizer2.Add(self.m_panel1, 1, wx.EXPAND | wx.ALL, 5)

        bSizer1.Add(bSizer2, 1, wx.EXPAND, 5)

        self.SetSizer(bSizer1)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.button_login.Bind(wx.EVT_BUTTON, self.get_project_list)
        self.button_go.Bind(wx.EVT_BUTTON, self.get_data_all)
        self.button_close.Bind(wx.EVT_BUTTON, self.windows_close)

        self._thread = Thread(target=self.run, args=())
        self._thread.daemon = True

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class

    def windows_close(self, event):
        self.Close()

    def get_data_all(self, event):
        self._thread.start()
        self.started = True
        self.button_go = event.GetEventObject()
        self.button_go.Disable()

    def get_project_list(self, event):
        username = self.input_username.GetValue().strip()
        password = self.input_password.GetValue().strip()
        list_project_name = []
        # driverpath = os.path.join(os.path.abspath(os.path.curdir), "chromedriver.exe")
        # browser = webdriver.Chrome(driverpath)
        driverpath = os.path.join(os.path.abspath(os.path.curdir), "phantomjs.exe")
        browser = webdriver.PhantomJS(driverpath)
        url = "http://10.7.13.21:2000/"
        browser.get(url)
        browser.find_element_by_id("userName").send_keys(username)
        browser.find_element_by_id("userPassword").send_keys(password)
        browser.find_element_by_css_selector("#loginBtn").click()
        time.sleep(5)
        browser.find_element_by_css_selector("#rdmLeft > ul > li:nth-child(5) > div.select-top-menu > a")
        ActionChains(browser).move_to_element(
            browser.find_element_by_css_selector("#rdmLeft > ul > li:nth-child(5) > div.select-top-menu > a")).perform()
        browser.find_element_by_css_selector(
            "#rdmLeft > ul > li:nth-child(5) > div.select-top-menu > a").click()
        time.sleep(2)
        WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, 'iframe#mainFrame')))
        browser.switch_to.frame('mainFrame')
        all_project_link = browser.find_elements_by_css_selector("#bodyPanel > table:nth-child(1) > tbody > tr")
        for item in all_project_link:
            project_name_text = item.find_element_by_xpath("td[4]/div/span/a").text
            list_project_name.append(project_name_text)
        self.combox_project_list.Set(list_project_name)
        browser.quit()
        dlg_info = wx.MessageDialog(None, '项目列表已经获取完毕！'.decode('gbk'), '完成提示'.decode('gbk'), wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        dlg_info.ShowModal()

    def run(self):
        self.button_go.Disable()
        self.updatedisplay(("开始抓取信息，开始时间{time_start}".format(time_start=time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))).decode('gbk'))
        username = self.input_username.GetValue().strip()
        password = self.input_password.GetValue().strip()
        project_name_selected = self.combox_project_list.GetValue()
        list_project_name = []
        driverpath = os.path.join(os.path.abspath(os.path.curdir), "chromedriver.exe")
        browser = webdriver.Chrome(driverpath)
        # firefoxdriverPath = os.path.abspath(os.path.curdir)
        # browser = webdriver.Firefox(firefoxdriverPath)
        browser.maximize_window()
        time.sleep(2)
        # driverpath = os.path.join(os.path.abspath(os.path.curdir), "phantomjs.exe")
        # browser = webdriver.PhantomJS(driverpath)
        #进入登录界面
        url = "http://10.7.13.21:2000/"
        browser.get(url)
        time.sleep(1)
        browser.find_element_by_id("userName").send_keys(username)
        browser.find_element_by_id("userPassword").send_keys(password)
        browser.find_element_by_css_selector("#loginBtn").click()
        time.sleep(2)
        browser.find_element_by_css_selector("div#rdmLeft > ul > li:nth-child(5) > div.select-top-menu > a")
        ActionChains(browser).move_to_element(
            browser.find_element_by_css_selector("div#rdmLeft > ul > li:nth-child(5) > div.select-top-menu > a")).perform()
        browser.find_element_by_css_selector(
            "div#rdmLeft > ul > li:nth-child(5) > div.select-top-menu > a").click()
        WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, 'iframe#mainFrame')))
        browser.switch_to.frame('mainFrame')
        time.sleep(0.5)
        ActionChains(browser).move_to_element(browser.find_element_by_link_text("{projectname}".format(projectname=project_name_selected))).perform()
        browser.find_element_by_link_text("{projectname}".format(projectname=project_name_selected)).click()
        WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, 'iframe#entityTab_')))
        browser.switch_to.frame('entityTab_')
        browser.find_element_by_css_selector("li#li_ISU > div").click()
        WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, 'iframe#tabs_panel_9')))
        browser.switch_to.frame('tabs_panel_9')
        time.sleep(0.5)
        list_sn = []
        list_issue_type = []
        list_title = []
        list_operation_people = []
        list_create_date = []
        list_verify_date = []
        total_pages = int(browser.find_element_by_css_selector("span#page_count").text.strip())
        while True:
            list_issue = browser.find_elements_by_css_selector("div#bodyPanel > table:nth-child(1) > tbody > tr")
            length = len(list_issue)
            for count_line in range(1, length + 1):
                sn_link = browser.find_element_by_css_selector("div#bodyPanel > table:nth-child(1) > tbody > tr:nth-child(%d) > td:nth-child(2) > div" % count_line)
                sn = sn_link.text.strip()
                link_title = browser.find_element_by_xpath('//div[@id="bodyPanel"]/table[1]/tbody/tr[%d]//td[4]/div/span/a' % count_line)
                create_date_temp = browser.find_element_by_css_selector("div#bodyPanel > table:nth-child(1) > tbody > tr:nth-child(%d) > td:nth-child(10) > div" % count_line).text.strip().split(" ")[0].split("-")
                create_date = "/".join(create_date_temp)
                issue_type = browser.find_element_by_css_selector("div#bodyPanel > table:nth-child(1) > tbody > tr:nth-child(%d) > td:nth-child(3) > div" % count_line).text.strip()
                title = link_title.text.strip()
                ActionChains(browser).move_to_element(sn_link).double_click().perform()
                time.sleep(1)
                browser.switch_to.parent_frame()
                browser.switch_to.frame('workflowFrame')
                ActionChains(browser).move_to_element(browser.find_element_by_css_selector("div#tabPane > div > ul > li:nth-child(4) > div")).perform()
                browser.find_element_by_css_selector("div#tabPane > div > ul > li:nth-child(4) > div").click()
                time.sleep(0.5)
                browser.switch_to.frame('tabs_panel_3')
                browser.switch_to.frame('viewFrame')
                list_status = browser.find_elements_by_css_selector("table#procTab > tbody > tr")
                length_status = len(list_status)
                for count_status in range(2, length_status):
                    status = browser.find_element_by_css_selector("table#procTab > tbody > tr:nth-child(%d) > td:nth-child(2) > div" % count_status).text.strip()
                    operation = browser.find_element_by_css_selector("table#procTab > tbody > tr:nth-child(%d) > td:nth-child(3) > div" % count_status).text.strip()
                    operation_date_temp = browser.find_element_by_css_selector("table#procTab > tbody > tr:nth-child(%d) > td:nth-child(5) > div" % count_status).text.strip().split(" ")[0].split("-")
                    operation_date = "/".join(operation_date_temp)
                    operation_people = browser.find_element_by_css_selector("table#procTab > tbody > tr:nth-child(%d) > td:nth-child(4) > div" % count_status).text.strip()
                    if status == 'Verify' and operation == '驳回'.decode('gbk'):
                        list_sn.append(sn)
                        self.updatedisplay(sn)
                        list_issue_type.append(issue_type)
                        list_title.append(title)
                        list_operation_people.append(operation_people)
                        list_create_date.append(create_date)
                        list_verify_date.append(operation_date)
                        break
                browser.switch_to.parent_frame()
                browser.switch_to.parent_frame()
                ActionChains(browser).move_to_element(browser.find_element_by_css_selector("a#returnLink")).perform()
                browser.find_element_by_css_selector("a#returnLink").click()
                time.sleep(0.5)
                browser.switch_to.default_content()
                WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, 'iframe#mainFrame')))
                browser.switch_to.frame('mainFrame')
                WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, 'iframe#entityTab_')))
                browser.switch_to.frame('entityTab_')
                WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, 'iframe#tabs_panel_9')))
                browser.switch_to.frame('tabs_panel_9')
            browser.switch_to.default_content()
            WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, 'iframe#mainFrame')))
            browser.switch_to.frame('mainFrame')
            WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, 'iframe#entityTab_')))
            browser.switch_to.frame('entityTab_')
            WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, 'iframe#tabs_panel_9')))
            browser.switch_to.frame('tabs_panel_9')
            current_page_temp = browser.find_element_by_css_selector("#targetPage").get_attribute("Value")
            current_page = int(current_page_temp)
            self.updatedisplay(("完成{current_page_sub}/{total_page_sub}页信息抓取".format(current_page_sub=current_page, total_page_sub=total_pages)).decode('gbk'))
            if current_page == total_pages:
                browser.quit()
                break
            else:
                browser.find_element_by_css_selector("#pagination_nextPage > img").click()
                time.sleep(1)
        self.button_go.Enable()
        self.updatedisplay(("结束抓取信息，开始时间{time_end}".format(time_end=time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))).decode('gbk'))

        dlg_info = wx.MessageDialog(None, 'BUG一次解决率信息已经抓取完毕'.decode('gbk'), '完成提示'.decode('gbk'), wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        dlg_info.ShowModal()

        title_sheet = ['编号'.decode('gbk'), '所属类别'.decode('gbk'), '标题'.decode('gbk'),  '驳回人'.decode('gbk'),'创建时间'.decode('gbk'), '提交方案驳回时间'.decode('gbk')]
        timestamp = time.strftime('%Y%m%d', time.localtime())
        workbook_display = xlsxwriter.Workbook('%s_BUG一次解决率统计-%s.xlsx'.decode('gbk') % (project_name_selected, timestamp))
        sheet = workbook_display.add_worksheet('BUG一次解决率统计'.decode('gbk'))
        formatone = workbook_display.add_format()
        formatone.set_border(1)
        formattwo = workbook_display.add_format()
        formattwo.set_border(1)
        formattitle = workbook_display.add_format()
        formattitle.set_border(1)
        formattitle.set_align('center')
        formattitle.set_bg_color("yellow")
        formattitle.set_bold(True)
        sheet.merge_range(0, 0, 0, 4, "%s_BUG一次解决率统计".decode('gbk') % project_name_selected, formattitle)
        sheet.set_column('A:B', 17)
        sheet.set_column('C:C', 65)
        sheet.set_column('D:D', 15)
        sheet.set_column('E:F', 18)
        for index_title, item_title in enumerate(title_sheet):
            sheet.write(1, index_title, item_title, formatone)
        for index_data, item_data in enumerate(list_sn):
            sheet.write(2 + index_data, 0, item_data, formatone)
            sheet.write(2 + index_data, 1, list_issue_type[index_data], formatone)
            sheet.write(2 + index_data, 2, list_title[index_data], formatone)
            sheet.write(2 + index_data, 3, list_operation_people[index_data], formatone)
            sheet.write_datetime(2 + index_data, 4, datetime.datetime.strptime(list_create_date[index_data], '%Y/%m/%d'), workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
            sheet.write_datetime(2 + index_data, 5, datetime.datetime.strptime(list_verify_date[index_data], '%Y/%m/%d'), workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
        workbook_display.close()

    def updatedisplay(self, msg):
        t = msg
        self.output_info.AppendText("%s".decode('gbk') % t)
        self.output_info.AppendText(os.linesep)


if __name__ == '__main__':
    app = wx.App()
    frame = BugSolutionMoreThanOnce(None)
    frame.Show()
    app.MainLoop()
