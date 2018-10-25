#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import requests
import re
from bs4 import BeautifulSoup
import xlsxwriter
import os
import time
import datetime
from threading import Thread
import wx
import urllib2
import base64
import HTMLParser


class BugSolutionMoreThanOnce(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"BUG统计", pos=wx.DefaultPosition,
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

    def updatedisplay(self, msg):
        t = msg
        self.output_info.AppendText("%s".decode('gbk') % t)
        self.output_info.AppendText(os.linesep)

    def get_project_list(self, event):
        username = self.input_username.GetValue().strip()
        password_input = self.input_password.GetValue().strip()
        # parser = HTMLParser.HTMLParser()
        password = base64.b64encode(password_input)
        # use engine to  get cookie
        url_getsession = "http://10.7.13.21:2000/scripts/base64.js?V=V808R08M02sp06"
        get_data = requests.session()
        session = get_data.get(url_getsession).cookies.values()[0]
        # login 1
        url_login = "http://10.7.13.21:2000/dwr/call/plaincall/jSystemFilterBean.checkSessionCode.dwr"
        payload_login = {
            'callCount': "1",
            'page': "/",
            'httpSessionId': "{cookie_session}".format(cookie_session=session),
            'scriptSessionId': "750F49C3DA1731E1EF25170F88D6526D343",
            'c0-scriptName': "jSystemFilterBean",
            'c0-methodName': "checkSessionCode",
            'c0-id': "0",
            'c0-param0': "string:N",
            'c0-param1': "string:",
            'batchId': "1"
        }
        headers_login = {
            'accept': "*/*",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.9",
            'connection': "keep-alive",
            'host': "10.7.13.21:2000",
            'origin': "http://10.7.13.21:2000",
            'referer': "http://10.7.13.21:2000/",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36",
            'content-type': "text/plain"
        }
        get_data.post(url_login, data=payload_login, headers=headers_login)

        # check code(login to)
        url_checkcode = "http://10.7.13.21:2000/j_security_check"
        headers_checkcode = {
            'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.9",
            'Cache-Control': "max-age=0",
            'Connection': "keep-alive",
            'Content-Type': "application/x-www-form-urlencoded",
            'Host': "10.7.13.21:2000",
            'Origin': "http://10.7.13.21:2000",
            'Referer': "http://10.7.13.21:2000/",
            'Upgrade-Insecure-Requests': "1",
            'User-Agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36"
        }
        payload_checkcode = "isExpires=1&sessionIndex=&j_username={username_value}&j_password={password_value}&j_validatecode=&remember=on&BROWSER_VERSION=1&REMOTE_LANGUAGE=undefined".format(
            username_value=username, password_value=password)
        get_data.post(url_checkcode, data=payload_checkcode, headers=headers_checkcode)
        # get post data for list projects
        url_page_list_data = "http://10.7.13.21:2000/pages/lifecycle/entity/list.jsf"
        querystring = {"type": "PJT"}
        headers = {
            'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.9",
            'connection': "keep-alive",
            'content-type': "application/x-www-form-urlencoded; charset=UTF-8",
            'origin': "http://10.7.13.21:2000",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36",
            'host': "10.7.13.21:2000",
            'referer': "http://10.7.13.21:2000/main.do",
            'upgrade-insecure-requests': "1",
            'cache-control': "no-cache"
        }
        page_content = get_data.get(url_page_list_data, headers=headers, params=querystring).text
        data_to_filter = BeautifulSoup(page_content, "html.parser")
        viewid = data_to_filter.select('li[issystem="Y"]')[0].attrs['val']
        # print data_to_filter
        javax_faces_ViewState_temp = data_to_filter.select('input[id="javax.faces.ViewState"]')[0].attrs['value']
        javax_faces_ViewState = urllib2.quote(javax_faces_ViewState_temp)
        cate = data_to_filter.select('input[name="cate"]')[0].attrs['value']

        # get project info
        url_list_project = "http://10.7.13.21:2000/pages/lifecycle/entity/list.jsf"
        headers_list_project = {
            'accept': "*/*",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.9",
            'connection': "keep-alive",
            'content-type': "application/x-www-form-urlencoded",
            'origin': "http://10.7.13.21:2000",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36",
            'host': "10.7.13.21:2000",
            'referer': "http://10.7.13.21:2000/pages/lifecycle/entity/list.jsf?type=PJT",
            'Cache-Control': "no-cache"
        }
        payload_list_project = "AJAXREQUEST=j_id_jsp_1770895026_0&operate=operate&module=PJT&hideNodes=&toggleNode=&hideAll=1&cate={catevalue}&viewId={viewidvalue}&targetPage=1&orderBy=&orderByType=&headCondition=&filterIds=&operate%3A_fileInfo=&javax.faces.ViewState={javavalue}&operate%3AreflushAll=operate%3AreflushAll".format(
            catevalue=cate, viewidvalue=viewid, javavalue=javax_faces_ViewState)
        data_list_project_temp = get_data.post(url_list_project, data=payload_list_project,
                                               headers=headers_list_project).text
        data_list_project = BeautifulSoup(data_list_project_temp, "html.parser")
        data_list_project_after_filter = data_list_project.select('span[id="_ajax:data"]')[0].text
        name_project = re.findall(r'\\"Name\\":{\\"v\\":\\"(.*?)\\",', data_list_project_after_filter)
        # belongid_project_list = re.findall(r'\\"ID\\":{\\"v\\":\\"(.*?)\\",', data_list_project_after_filter)
        name_project_list = [item.decode('unicode_escape') for item in name_project]
        self.combox_project_list.Set(name_project_list)
        dlg_info = wx.MessageDialog(None, '项目列表已经获取完毕！'.decode('gbk'), '完成提示'.decode('gbk'), wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        dlg_info.ShowModal()

    def run(self):
        self.updatedisplay(("开始抓取信息，开始时间{time_start}".format(time_start=time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))).decode('gbk'))
        project_name_selected = unicode(self.combox_project_list.GetValue())
        username = self.input_username.GetValue().strip()
        password_input = self.input_password.GetValue().strip()
        # parser = HTMLParser.HTMLParser()
        password = base64.b64encode(password_input)
        # use engine to  get cookie
        url_getsession = "http://10.7.13.21:2000/scripts/base64.js?V=V808R08M02sp06"
        get_data = requests.session()
        session = get_data.get(url_getsession).cookies.values()[0]
        # login 1
        url_login = "http://10.7.13.21:2000/dwr/call/plaincall/jSystemFilterBean.checkSessionCode.dwr"
        payload_login = {
            'callCount': "1",
            'page': "/",
            'httpSessionId': "{cookie_session}".format(cookie_session=session),
            'scriptSessionId': "750F49C3DA1731E1EF25170F88D6526D343",
            'c0-scriptName': "jSystemFilterBean",
            'c0-methodName': "checkSessionCode",
            'c0-id': "0",
            'c0-param0': "string:N",
            'c0-param1': "string:",
            'batchId': "1"
        }
        headers_login = {
            'accept': "*/*",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.9",
            'connection': "keep-alive",
            'host': "10.7.13.21:2000",
            'origin': "http://10.7.13.21:2000",
            'referer': "http://10.7.13.21:2000/",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36",
            'content-type': "text/plain"
        }
        get_data.post(url_login, data=payload_login, headers=headers_login)

        # check code(login to)
        url_checkcode = "http://10.7.13.21:2000/j_security_check"
        headers_checkcode = {
            'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.9",
            'Cache-Control': "max-age=0",
            'Connection': "keep-alive",
            'Content-Type': "application/x-www-form-urlencoded",
            'Host': "10.7.13.21:2000",
            'Origin': "http://10.7.13.21:2000",
            'Referer': "http://10.7.13.21:2000/",
            'Upgrade-Insecure-Requests': "1",
            'User-Agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36"
        }
        payload_checkcode = "isExpires=1&sessionIndex=&j_username={username_value}&j_password={password_value}&j_validatecode=&remember=on&BROWSER_VERSION=1&REMOTE_LANGUAGE=undefined".format(username_value=username, password_value=password)
        get_data.post(url_checkcode, data=payload_checkcode, headers=headers_checkcode)
        # get post data for list projects
        url_page_list_data = "http://10.7.13.21:2000/pages/lifecycle/entity/list.jsf"
        querystring = {"type": "PJT"}
        headers = {
            'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.9",
            'connection': "keep-alive",
            'content-type': "application/x-www-form-urlencoded; charset=UTF-8",
            'origin': "http://10.7.13.21:2000",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36",
            'host': "10.7.13.21:2000",
            'referer': "http://10.7.13.21:2000/main.do",
            'upgrade-insecure-requests': "1",
            'cache-control': "no-cache"
        }
        page_content = get_data.get(url_page_list_data, headers=headers, params=querystring).text
        data_to_filter = BeautifulSoup(page_content, "html.parser")
        viewid = data_to_filter.select('li[issystem="Y"]')[0].attrs['val']
        # print data_to_filter
        javax_faces_ViewState_temp = data_to_filter.select('input[id="javax.faces.ViewState"]')[0].attrs['value']
        javax_faces_ViewState = urllib2.quote(javax_faces_ViewState_temp)
        cate = data_to_filter.select('input[name="cate"]')[0].attrs['value']

        # get project info
        self.updatedisplay("开始获取所有项目信息".decode('gbk'))
        url_list_project = "http://10.7.13.21:2000/pages/lifecycle/entity/list.jsf"
        headers_list_project = {
            'accept': "*/*",
            'accept-encoding': "gzip, deflate",
            'accept-language': "zh-CN,zh;q=0.9",
            'connection': "keep-alive",
            'content-type': "application/x-www-form-urlencoded",
            'origin': "http://10.7.13.21:2000",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36",
            'host': "10.7.13.21:2000",
            'referer': "http://10.7.13.21:2000/pages/lifecycle/entity/list.jsf?type=PJT",
            'Cache-Control': "no-cache"
        }
        payload_list_project = "AJAXREQUEST=j_id_jsp_1770895026_0&operate=operate&module=PJT&hideNodes=&toggleNode=&hideAll=1&cate={catevalue}&viewId={viewidvalue}&targetPage=1&orderBy=&orderByType=&headCondition=&filterIds=&operate%3A_fileInfo=&javax.faces.ViewState={javavalue}&operate%3AreflushAll=operate%3AreflushAll".format(catevalue=cate, viewidvalue=viewid, javavalue=javax_faces_ViewState)
        data_list_project_temp = get_data.post(url_list_project, data=payload_list_project, headers=headers_list_project).text
        data_list_project = BeautifulSoup(data_list_project_temp, "html.parser")
        data_list_project_after_filter = data_list_project.select('span[id="_ajax:data"]')[0].text
        name_project = re.findall(r'\\"Name\\":{\\"v\\":\\"(.*?)\\",', data_list_project_after_filter)
        belongid_project_list = re.findall(r'\\"ID\\":{\\"v\\":\\"(.*?)\\",', data_list_project_after_filter)
        name_project_list = [item.decode('unicode_escape') for item in name_project]
        belongid = belongid_project_list[name_project_list.index(project_name_selected)]

        # get bug list for one project
        # preparetion for buglist
        url_buglist = "http://10.7.13.21:2000/pages/entity/belongList.jsf"
        querystring_buglist_pre = {"belongId": "{belongid_value}".format(belongid_value=belongid), "belongType": "PJT", "type": "ISU"}
        headers_buglist_pre = {
            'Accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            'Accept-Encoding': "gzip, deflate",
            'Accept-Language': "zh-CN,zh;q=0.9",
            'Connection': "keep-alive",
            'Host': "10.7.13.21:2000",
            'Referer': "http://10.7.13.21:2000/pages/lifecycle/entity/entityTab.project.jsf?objectId={belongid_value}&lcType=PJT&r=0.6253874309481038".format(belongid_value=belongid),
            'Upgrade-Insecure-Requests': "1",
            'User-Agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36",
            'Cache-Control': "no-cache"
        }
        data_buglist_pre_temp = get_data.get(url_buglist, headers=headers_buglist_pre, params=querystring_buglist_pre).text
        data_buglist_pre = BeautifulSoup(data_buglist_pre_temp, "html.parser")
        viewid_buglist = data_buglist_pre.select('input[id="viewId"]')[0].attrs["value"]
        viewstate_buglist = urllib2.quote(data_buglist_pre.select('input[id="javax.faces.ViewState"]')[0].attrs["value"])

        # get total pages
        self.updatedisplay("开始获取所有BUG的总页数".decode('gbk'))
        headers_buglist = {
            'Content-Type': "application/x-www-form-urlencoded; charset=UTF-8",
            'Accept': "*/*",
            'Accept-Encoding': "gzip, deflate",
            'Accept-Language': "zh-CN,zh;q=0.9",
            'Connection': "keep-alive",
            'Host': "10.7.13.21:2000",
            'Origin': "http://10.7.13.21:2000",
            'Cache-Control': "no-cache",
            'User-Agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36",
            'Referer': "http://10.7.13.21:2000/pages/entity/belongList.jsf?belongId={belongid_value}&belongType=PJT&type=ISU".format(belongid_value=belongid)
        }
        payload_buglist_pageone = "AJAXREQUEST=j_id_jsp_135821430_0&operate=operate&module=ISU&hideNodes=&toggleNode=&hideAll=1&cate=4&viewId={viewid_value}&targetPage=1&orderBy=&orderByType=&belongId={belongid_value}&belongType=PJT&headCondition=&filterIds=&operate%3A_fileInfo=&javax.faces.ViewState={viewstate_value}&operate%3AreflushAll=operate%3AreflushAll".format(viewid_value=viewid_buglist, belongid_value=belongid, viewstate_value=viewstate_buglist)
        buglist_page_one = BeautifulSoup(get_data.post(url_buglist, data=payload_buglist_pageone, headers=headers_buglist).text, "html.parser")
        total_pages = buglist_page_one.select('span[id="page_count"]')[0].text
        total_pages_count = int(total_pages) + 1
        self.updatedisplay("所有BUG的总页数为%s".decode('gbk') % str(total_pages_count))

        bug_name_list_total_temp = []
        bug_id_list_total_temp = []
        bug_status_list_total_temp = []
        bug_type_list_total_temp = []
        bug_create_date_list_total_temp = []
        bug_modify_date_list_total_temp = []
        bug_stage_list_total_temp = []
        bug_person_commit_list_temp = []

        # 从最开始的汇总页面收集BUG的信息，包括BUG的名字、当前状态、id编号、类别、创建时间、最后修改时间、什么阶段的等
        self.updatedisplay("开始分页获取所有BUG的总信息".decode('gbk'))
        for count in range(1, total_pages_count):
            self.updatedisplay("获取第%s页的BUG".decode('gbk') % str(count))
            payload_buglist = "AJAXREQUEST=j_id_jsp_135821430_0&operate=operate&module=ISU&hideNodes=&toggleNode=&hideAll=1&cate=4&viewId={viewid_value}&targetPage={page_value}&orderBy=&orderByType=&belongId={belongid_value}&belongType=PJT&headCondition=&filterIds=&operate%3A_fileInfo=&javax.faces.ViewState={viewstate_value}&operate%3AreflushAll=operate%3AreflushAll".format(viewid_value=viewid_buglist, belongid_value=belongid, viewstate_value=viewstate_buglist, page_value=count)
            buglist_for_one_page_temp = BeautifulSoup(get_data.post(url_buglist, data=payload_buglist, headers=headers_buglist).text, "html.parser")
            buglist_for_one_page = buglist_for_one_page_temp.select('span[id="_ajax:data"]')[0].text
            # print(buglist_for_one_page)
            # bug的名字
            bug_name_temp = re.findall(r'\\"Name\\":{\\"v\\":\\"(.*?)\\",', buglist_for_one_page)# [Apollo PVT\x5D[CMC 3.0.0\x5D\u901A\u8FC7IPMI\u547D\u4EE4\u5E26\u5916\u8BBE\u7F6ECMC\u7F51\u7EDC\u9759\u6001IP\uFF0C\u547D\u4EE4\u62A5\u9519\u4F46IP\u53EF\u4EE5\u767B\u5F55CMC
            bug_name = [item.decode('unicode_escape') for item in bug_name_temp]
            # bug提出人
            bug_person_commit_temp = re.findall(r'\\"CreatorID\\":{\\"v\\":\\".*?\\",\\"t\\":\\"(.*?)\\"},', buglist_for_one_page)
            bug_person_commit = [item.decode('unicode_escape') for item in bug_person_commit_temp]
            # bug当前的状态
            bug_status_temp = re.findall(r'\\"StatusID\\":{\\"v\\":\\".*?\\",\\"t\\":\\"(.*?)\\"}', buglist_for_one_page)# Verify
            bug_status = [item.decode('unicode_escape') for item in bug_status_temp]
            # bug的id编号
            bug_id = re.findall(r'\\"ID\\":{\\"v\\":\\"(.*?)\\",', buglist_for_one_page)# 99547a7e-7130-4821-a087-e71b01f0f05e
            # bug的类别，如HW，BMC，BIOS
            bug_type_temp = re.findall(r',\\"TypeID\\":{\\"v\\":\\".*?\\",\\"t\\":\\"(.*?)\\"', buglist_for_one_page)
            bug_type = [item.decode('unicode_escape') for item in bug_type_temp]
            # 创建时间
            bug_create_date = re.findall(r',\\"CreatedTime\\":{\\"v\\":\\".*?\\",\\"t\\":\\"(.*?)\\"},', buglist_for_one_page)
            # 最后修改时间
            bug_modify_date = re.findall(r':{\\"ModifyTime\\":{\\"v\\":\\".*?\\",\\"t\\":\\"(\d+-\d+-\d+).*?\\"', buglist_for_one_page)
            # stage
            bug_stage_temp = re.findall(r',\\"Fld_S_00020\\":{\\"v\\":\\".*?\\",\\"t\\":\\"(.*?)\\"},', buglist_for_one_page)
            bug_stage = [item.decode('unicode_escape') for item in bug_stage_temp]

            bug_name_list_total_temp.extend(bug_name)
            bug_status_list_total_temp.extend(bug_status)
            bug_id_list_total_temp.extend(bug_id)
            bug_type_list_total_temp.extend(bug_type)
            bug_create_date_list_total_temp.extend(bug_create_date)
            bug_modify_date_list_total_temp.extend(bug_modify_date)
            bug_stage_list_total_temp.extend(bug_stage)
            bug_person_commit_list_temp.extend(bug_person_commit)

        print len(bug_name_list_total_temp)
        print len(bug_status_list_total_temp)
        print len(bug_id_list_total_temp)
        print len(bug_create_date_list_total_temp)
        print len(bug_modify_date_list_total_temp)
        print len(bug_stage_list_total_temp)
        print len(bug_person_commit_list_temp)

        # 去除掉总体状态不正确，导致不显示的；
        bug_name_list_total = []
        bug_id_list_total = []
        bug_status_list_total = []
        bug_type_list_total = []
        bug_create_date_list_total = []
        bug_modify_date_list_total = []
        bug_stage_list_total = []
        bug_person_commit_list_total = []

        for index_statusid, item_statusid in enumerate(bug_status_list_total_temp):
            if len(item_statusid) != 0:
                # print(index_statusid)
                bug_name_list_total.append(bug_name_list_total_temp[index_statusid])
                bug_id_list_total.append(bug_id_list_total_temp[index_statusid])
                bug_status_list_total.append(item_statusid)
                bug_type_list_total.append(bug_type_list_total_temp[index_statusid])
                bug_create_date_list_total.append(bug_create_date_list_total_temp[index_statusid])
                bug_modify_date_list_total.append(bug_modify_date_list_total_temp[index_statusid])
                bug_stage_list_total.append(bug_stage_list_total_temp[index_statusid])
                bug_person_commit_list_total.append(bug_person_commit_list_temp[index_statusid])

        url_bug_first_page = []
        # bug_creatime_list_total = []
        bug_management_sn_list_total = []
        bug_description_list_total = []
        bug_solution_list_total = []
        bug_rootcause_list_total = []

        # get link address for  every bug
        self.updatedisplay("开始获取所有BUG的链接信息".decode('gbk'))
        for index_bug_id, item_bug_id in enumerate(bug_id_list_total):
            url_bug_detail = "http://10.7.13.21:2000/pages/workflow/entityTab.jsf"
            querystring_bug_detail = {"workflowType": "ISU", "objectId": "{bugid_value}".format(bugid_value=item_bug_id)}
            data_bug_detail = BeautifulSoup(get_data.get(url_bug_detail, params=querystring_bug_detail).text, "html.parser")
            url_bug_first_page_temp = data_bug_detail.select('div[id="tabPane"] > ul:nth-of-type(1) > li:nth-of-type(1)')[0].attrs["url"]
            url_bug_first_page.append("http://10.7.13.21:2000" + url_bug_first_page_temp)

        self.updatedisplay("开始获取所有BUG的首页信息".decode('gbk'))
        for index_bug_first_page, item_bug_first_page in enumerate(url_bug_first_page):
            # print bug_status_list_total[index_bug_first_page]
            bug_detail_first_page = BeautifulSoup(get_data.get(item_bug_first_page).text, "html.parser")
            # print bug_detail_first_page
            data_detail_first_page_temp = bug_detail_first_page.select('input[id="data"]')[0].attrs["value"]
            # print data_detail_first_page_temp
            management_sn = re.findall(',"b":"编号","v":"(.*?)",'.decode('gbk'), data_detail_first_page_temp)[0]
            try:
                bug_description = re.findall(',"b":"问题描述","v":"(.*?)",'.decode('gbk'), data_detail_first_page_temp)[0]
            except IndexError:
                bug_description = "None"
            try:
                bug_solution = re.findall(',"b":"解决方案","v":"(.*?)",'.decode('gbk'), data_detail_first_page_temp)[0]
            except IndexError:
                bug_solution = "None"
            try:
                bug_rootcause = re.findall(',"b":"Root cause 确认","v":"(.*?)",'.decode('gbk'), data_detail_first_page_temp)[0]
            except IndexError:
                bug_rootcause = "None"
            # #创建时间优先从页面元素获取。如果不存在（BUG已关闭），再使用下方的评论时间。
            # createtime_temp = re.findall(',"b":"提出日期","v":"(.*?)",'.decode('gbk'), data_detail_first_page_temp)
            # if len(createtime_temp) == 0 :
            #     createtime = bug_detail_first_page.select('#noteDiv > div.notes > div.title')[-1].text.strip().split(",")[-1].strip().split(" ")[0].strip()
            # else:
            #     #print createtime_temp
            #     createtime = createtime_temp[0]
            # bug_creatime_list_total.append(createtime)
            bug_management_sn_list_total.append(management_sn)
            bug_description_list_total.append(bug_description)
            bug_solution_list_total.append(bug_solution)
            bug_rootcause_list_total.append(bug_rootcause)

        bug_operation_time_list = []
        bug_refused_or_not_list = []
        bug_username_operation_list = []
        url_workflow_list = []

        # 获取记录BUG操作流程的页面链接，获取驳回的次数、驳回人列表、驳回的时间列表
        self.updatedisplay("开始获取所有BUG的操作流程".decode('gbk'))
        for index_bug_id, item_bug_id in enumerate(bug_id_list_total):
            url_bug_workflow = "http://10.7.13.21:2000/pages/workflow/workflowChartByObject.jsf"
            querystring_bug_workflow = {"workflowType": "ISU", "objectId": "{bugid_value}".format(bugid_value=item_bug_id)}
            bug_workflow = BeautifulSoup(get_data.get(url_bug_workflow, params=querystring_bug_workflow).text, "html.parser")
            url_workflow_temp = bug_workflow.select('iframe[id="viewFrame"]')[0].attrs['src']
            url_workflow = "http://10.7.13.21:2000" + url_workflow_temp.replace("../../", "/")
            url_workflow_list.append(url_workflow)

        # 根据上一步的连接获取BUG操作记录
        bug_operation_date_list = []
        for index_operation, item_operation in enumerate(url_workflow_list):
            bug_refuse_list_temp = []
            bug_username_operation_list_temp = []
            bug_operation_time_list_temp = []
            bug_operation_date_list_temp = []
            data_workflow_page = BeautifulSoup(get_data.get(item_operation).text, "html.parser")
            status_operation_list_temp = data_workflow_page.select('table[id="procTab"] > tr > td:nth-of-type(2) > div')
            name_operation_list_temp = data_workflow_page.select('table[id="procTab"] > tr > td:nth-of-type(3) > div')
            person_operation_list_temp = data_workflow_page.select('table[id="procTab"] > tr > td:nth-of-type(4) > div')
            time_operation_list_temp = data_workflow_page.select('table[id="procTab"] > tr > td:nth-of-type(5) > div')
            for index_operation_status, item_operation_status in enumerate(status_operation_list_temp):
                # 如果出现被驳回，则记录是否被驳回，然后记录驳回次数、驳回人列表、驳回时间列表；如果没有出现过被驳回，则驳回次数为0，其他两个列表都是空。
                if item_operation_status.text.strip() == "Verify" and (name_operation_list_temp[index_operation_status].text.strip() == "驳回".decode('gbk') or name_operation_list_temp[index_operation_status].text.strip() == "验证不通过".decode('gbk') ):
                    bug_username_operation = person_operation_list_temp[index_operation_status].text.strip()
                    bug_operation_time_temp = time_operation_list_temp[index_operation_status].text.strip()
                    bug_operation_time = bug_operation_time_temp.split(" ")[0].strip()

                    bug_refuse_list_temp.append("1")
                    bug_username_operation_list_temp.append(bug_username_operation)
                    bug_operation_time_list_temp.append(bug_operation_time)
                    bug_operation_date_list_temp.append(bug_operation_time_temp)
            bug_operation_date_list.append(bug_operation_date_list_temp)
            bug_username_operation_list.append(bug_username_operation_list_temp)
            bug_operation_time_list.append(bug_operation_time_list_temp)
            bug_refused_or_not_list.append(len(bug_refuse_list_temp))

        # 获取记录BUG操作记录的页面链接,来获取驳回的记录的原因
        bug_refuse_reason_list = []
        operation_status_list = ["驳回".decode('gbk'), "验证不通过".decode('gbk')]
        self.updatedisplay("开始获取所有BUG的操作记录".decode('gbk'))
        for index_bug_id, item_bug_id in enumerate(bug_id_list_total):
            detail_refuse_data_list = []
            url_bug_operation = "http://10.7.13.21:2000/pages/workflow/operateRecordList.jsf"
            querystring_bug_operation = {"workflowType": "ISU", "objectId": "{bugid_value}".format(bugid_value=item_bug_id)}
            bug_operation_detail = BeautifulSoup(get_data.get(url_bug_operation, params=querystring_bug_operation).text, "html.parser")
            data_detail_operation_list_temp = bug_operation_detail.select('div[id="recordDiv"] > div > span:nth-of-type(2)')
            data_detail_operation_list_temp_temp = []
            # 排除评论的条目
            for item_operation_temp_temp in data_detail_operation_list_temp:
                temp_1 = item_operation_temp_temp.text.strip()
                if temp_1 != "评论".decode('gbk'):
                    data_detail_operation_list_temp_temp.append(item_operation_temp_temp)
            # 获取驳回时的评论信息
            try:
                for index_operation_temp, item_operation_temp in enumerate(data_detail_operation_list_temp_temp):
                    temp_1 = item_operation_temp.text.strip()
                    temp_2 = item_operation_temp.find_parent("div").text.split(",")[-1].strip()
                    if temp_2 in bug_operation_date_list[index_bug_id] and temp_1 in operation_status_list:
                        data_temp = item_operation_temp.find_parent("div").find_next_sibling("div").text.strip()
                        if data_temp is not None and len(data_temp) != 0:
                            detail_refuse_data_list.append(data_temp)
                            print(data_temp)
                            print("#############")
            except IndexError:
                pass
            bug_refuse_reason_list.append(";".join(detail_refuse_data_list))

        # write to log file
        title_sheet = ['编号'.decode('gbk'), '类别'.decode('gbk'), '标题'.decode('gbk'), '当前状态'.decode('gbk'), 'BUG创建时间'.decode('gbk'), '提出人'.decode('gbk'), '最后更新时间'.decode('gbk'), '问题发现阶段'.decode('gbk'), '方案被驳回次数'.decode('gbk'), '驳回人列表'.decode('gbk'), '提交方案驳回时间列表'.decode('gbk'), '驳回原因列表'.decode('gbk'), '问题描述'.decode('gbk'), '解决方案'.decode('gbk'), 'Root Cause'.decode('gbk')]
        timestamp = time.strftime('%Y%m%d', time.localtime())
        workbook_display = xlsxwriter.Workbook('%s_BUG统计-%s.xlsx'.decode('gbk') % (project_name_selected, timestamp))
        sheet = workbook_display.add_worksheet('%s_BUG统计'.decode('gbk') % project_name_selected)
        formatone = workbook_display.add_format()
        formatone.set_border(1)
        formattwo = workbook_display.add_format()
        formattwo.set_border(1)
        formattitle = workbook_display.add_format()
        formattitle.set_border(1)
        formattitle.set_align('center')
        formattitle.set_bg_color("yellow")
        formattitle.set_bold(True)
        sheet.merge_range(0, 0, 0, 14, "%s_BUG情况统计".decode('gbk') % project_name_selected, formattitle)
        sheet.set_column('A:A', 17)
        sheet.set_column('B:B', 10)
        sheet.set_column('C:C', 55)
        sheet.set_column('D:G', 15)
        sheet.set_column('H:H', 11)
        sheet.set_column('I:I', 16)
        sheet.set_column('J:J', 11)
        sheet.set_column('K:K', 18)
        sheet.set_column('L:O', 16)

        # sheet.set_column('E:F', 18)
        for index_title, item_title in enumerate(title_sheet):
            sheet.write(1, index_title, item_title, formatone)
        for index_data, item_data in enumerate(bug_management_sn_list_total):
            # 编号
            sheet.write(2 + index_data, 0, item_data, formatone)
            # 类别
            sheet.write(2 + index_data, 1, bug_type_list_total[index_data], formatone)
            # BUG名字
            sheet.write(2 + index_data, 2, bug_name_list_total[index_data], formatone)
            # 当前状态
            sheet.write(2 + index_data, 3, bug_status_list_total[index_data], formatone)
            # BUG创建时间
            sheet.write_datetime(2 + index_data, 4, datetime.datetime.strptime(bug_create_date_list_total[index_data], '%Y-%m-%d'), workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
            # BUG提出人
            sheet.write(2 + index_data, 5, bug_person_commit_list_total[index_data], formatone)
            # 最后更新时间
            sheet.write_datetime(2 + index_data, 6, datetime.datetime.strptime(bug_modify_date_list_total[index_data], '%Y-%m-%d'), workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
            # 问题发现阶段
            sheet.write(2 + index_data, 7, bug_stage_list_total[index_data], formatone)
            # 方案被驳回次数
            sheet.write(2 + index_data, 8, bug_refused_or_not_list[index_data], formatone)
            # 驳回人
            if bug_refused_or_not_list[index_data] == 0:
                sheet.write(2 + index_data, 9, "None", formatone)
            else:
                sheet.write(2 + index_data, 9, ";".join(bug_username_operation_list[index_data]), formatone)
            # 驳回时间
            if bug_refused_or_not_list[index_data] == 0:
                sheet.write(2 + index_data, 10, "None", formatone)
            else:
                sheet.write(2 + index_data, 10, ";".join(bug_operation_time_list[index_data]), formatone)
            # 驳回原因
            sheet.write(2 + index_data, 11, bug_refuse_reason_list[index_data], formatone)
            # 问题描述
            sheet.write(2 + index_data, 12, bug_description_list_total[index_data], formatone)
            # 解决方案
            sheet.write(2 + index_data, 13, bug_solution_list_total[index_data], formatone)
            # root cause
            sheet.write(2 + index_data, 14, bug_rootcause_list_total[index_data], formatone)
        workbook_display.close()
        self.button_go.Enable()
        self.updatedisplay(("结束抓取信息，结束时间{time_end}".format(time_end=time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))).decode('gbk'))


if __name__ == '__main__':
    app = wx.App()
    frame = BugSolutionMoreThanOnce(None)
    frame.Show()
    app.MainLoop()
