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

    def updatedisplay(self, msg):
        t = msg
        self.output_info.AppendText("%s".decode('gbk') % t)
        self.output_info.AppendText(os.linesep)

    def get_project_list(self, event):
        username = self.input_username.GetValue().strip()
        password_input = self.input_password.GetValue().strip()
        parser = HTMLParser.HTMLParser()
        password=base64.b64encode(password_input)
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
        login = get_data.post(url_login, data=payload_login, headers=headers_login)

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
        checkcode = get_data.post(url_checkcode, data=payload_checkcode, headers=headers_checkcode)
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
        belongid_project_list = re.findall(r'\\"ID\\":{\\"v\\":\\"(.*?)\\",', data_list_project_after_filter)
        name_project_list = [item.decode('unicode_escape') for item in name_project]
        self.combox_project_list.Set(name_project_list)
        dlg_info = wx.MessageDialog(None, '项目列表已经获取完毕！'.decode('gbk'), '完成提示'.decode('gbk'), wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        dlg_info.ShowModal()

    def run(self):
        self.updatedisplay(("开始抓取信息，开始时间{time_start}".format(time_start=time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))).decode('gbk'))
        project_name_selected = unicode(self.combox_project_list.GetValue())
        username = self.input_username.GetValue().strip()
        password_input = self.input_password.GetValue().strip()
        parser = HTMLParser.HTMLParser()
        password=base64.b64encode(password_input)
        #use engine to  get cookie
        url_getsession = "http://10.7.13.21:2000/scripts/base64.js?V=V808R08M02sp06"
        get_data = requests.session()
        session = get_data.get(url_getsession).cookies.values()[0]
        #login 1
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
        login = get_data.post(url_login, data=payload_login, headers=headers_login)

        #check code(login to)
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
        checkcode = get_data.post(url_checkcode, data=payload_checkcode, headers=headers_checkcode)
        #get post data for list projects
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
        #print data_to_filter
        javax_faces_ViewState_temp = data_to_filter.select('input[id="javax.faces.ViewState"]')[0].attrs['value']
        javax_faces_ViewState=urllib2.quote(javax_faces_ViewState_temp)
        cate = data_to_filter.select('input[name="cate"]')[0].attrs['value']

        #get project info
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

        #get bug list for one project
        #preparetion for buglist
        url_buglist = "http://10.7.13.21:2000/pages/entity/belongList.jsf"
        querystring_buglist_pre = {"belongId":"{belongid_value}".format(belongid_value=belongid),"belongType":"PJT","type":"ISU"}
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

        #get total pages
        headers_buglist = {
            'Content-Type':"application/x-www-form-urlencoded; charset=UTF-8",
            'Accept':"*/*",
            'Accept-Encoding':"gzip, deflate",
            'Accept-Language':"zh-CN,zh;q=0.9",
            'Connection':"keep-alive",
            'Host':"10.7.13.21:2000",
            'Origin':"http://10.7.13.21:2000",
            'Cache-Control': "no-cache",
            'User-Agent':"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36",
            'Referer':"http://10.7.13.21:2000/pages/entity/belongList.jsf?belongId={belongid_value}&belongType=PJT&type=ISU".format(belongid_value=belongid)
        }
        payload_buglist_pageone = "AJAXREQUEST=j_id_jsp_135821430_0&operate=operate&module=ISU&hideNodes=&toggleNode=&hideAll=1&cate=4&viewId={viewid_value}&targetPage=1&orderBy=&orderByType=&belongId={belongid_value}&belongType=PJT&headCondition=&filterIds=&operate%3A_fileInfo=&javax.faces.ViewState={viewstate_value}&operate%3AreflushAll=operate%3AreflushAll".format(viewid_value=viewid_buglist, belongid_value=belongid, viewstate_value=viewstate_buglist)
        buglist_page_one = BeautifulSoup(get_data.post(url_buglist, data=payload_buglist_pageone, headers=headers_buglist).text, "html.parser")
        total_pages =  buglist_page_one.select('span[id="page_count"]')[0].text
        total_pages_count = int(total_pages) +1

        bug_name_list_total_temp = []
        bug_id_list_total_temp = []
        bug_status_list_total_temp = []
        #get detail for all bugs
        for count in range(1, total_pages_count):
            payload_buglist = "AJAXREQUEST=j_id_jsp_135821430_0&operate=operate&module=ISU&hideNodes=&toggleNode=&hideAll=1&cate=4&viewId={viewid_value}&targetPage={page_value}&orderBy=&orderByType=&belongId={belongid_value}&belongType=PJT&headCondition=&filterIds=&operate%3A_fileInfo=&javax.faces.ViewState={viewstate_value}&operate%3AreflushAll=operate%3AreflushAll".format(viewid_value=viewid_buglist, belongid_value=belongid, viewstate_value=viewstate_buglist, page_value=count)
            buglist_for_one_page_temp = BeautifulSoup(get_data.post(url_buglist, data=payload_buglist, headers=headers_buglist).text, "html.parser")
            buglist_for_one_page = buglist_for_one_page_temp.select('span[id="_ajax:data"]')[0].text
            bug_name_temp = re.findall(r'\\"Name\\":{\\"v\\":\\"(.*?)\\",', buglist_for_one_page)#[Apollo PVT\x5D[CMC 3.0.0\x5D\u901A\u8FC7IPMI\u547D\u4EE4\u5E26\u5916\u8BBE\u7F6ECMC\u7F51\u7EDC\u9759\u6001IP\uFF0C\u547D\u4EE4\u62A5\u9519\u4F46IP\u53EF\u4EE5\u767B\u5F55CMC
            bug_name = [item.decode('unicode_escape') for item in bug_name_temp]
            bug_status = re.findall(r'\\"StatusID\\":{\\"v\\":\\".*?\\",\\"t\\":\\"(.*?)\\"',buglist_for_one_page)#Verify
            bug_id = re.findall(r'\\"ID\\":{\\"v\\":\\"(.*?)\\",', buglist_for_one_page)#99547a7e-7130-4821-a087-e71b01f0f05e

            bug_name_list_total_temp.extend(bug_name)
            bug_status_list_total_temp.extend(bug_status)
            bug_id_list_total_temp.extend(bug_id)
        # print len(bug_name_list_total)
        # print len(bug_creatime_list_total)
        # print len(bug_management_sn_list_total)
        # print len(bug_category_list_total)
        # print len(bug_status_list_total)
        # print len(bug_id_list_total)
        #去除掉总体状态不正确，导致不显示的；
        bug_name_list_total = []
        bug_id_list_total = []
        bug_status_list_total = []
        for index_statusid, item_statusid in enumerate(bug_status_list_total_temp):
            if len(item_statusid) != 0:
                bug_name_list_total.append(bug_name_list_total_temp[index_statusid])
                bug_id_list_total.append(bug_id_list_total_temp[index_statusid])
                bug_status_list_total.append(item_statusid)


        url_bug_first_page = []
        bug_creatime_list_total = []
        bug_management_sn_list_total = []

        #get link address for  every bug
        for index_bug_id, item_bug_id in enumerate(bug_id_list_total):
            url_bug_detail = "http://10.7.13.21:2000/pages/workflow/entityTab.jsf"
            querystring_bug_detail = {"workflowType":"ISU","objectId":"{bugid_value}".format(bugid_value=item_bug_id)}
            data_bug_detail = BeautifulSoup(get_data.get(url_bug_detail, params=querystring_bug_detail).text, "html.parser")
            url_bug_first_page_temp = data_bug_detail.select('div[id="tabPane"] > ul:nth-of-type(1) > li:nth-of-type(1)')[0].attrs["url"]
            url_bug_first_page.append("http://10.7.13.21:2000" + url_bug_first_page_temp)
        for index_bug_first_page, item_bug_first_page in enumerate(url_bug_first_page):
            # print bug_status_list_total[index_bug_first_page]
            bug_detail_first_page = BeautifulSoup(get_data.get(item_bug_first_page).text, "html.parser")
            #print bug_detail_first_page
            data_detail_first_page_temp = bug_detail_first_page.select('input[id="data"]')[0].attrs["value"]
            #print data_detail_first_page_temp
            management_sn = re.findall(',"b":"编号","v":"(.*?)",'.decode('gbk'), data_detail_first_page_temp)[0]
            #创建时间优先从页面元素获取。如果不存在（BUG已关闭），再使用下方的评论时间。
            createtime_temp = re.findall(',"b":"提出日期","v":"(.*?)",'.decode('gbk'), data_detail_first_page_temp)
            if len(createtime_temp) == 0 :
                createtime = bug_detail_first_page.select('#noteDiv > div.notes > div.title')[-1].text.strip().split(",")[-1].strip().split(" ")[0].strip()
            else:
                #print createtime_temp
                createtime = createtime_temp[0]
            bug_management_sn_list_total.append(management_sn)
            bug_creatime_list_total.append(createtime)


        bug_name_list = []
        bug_creatime_list = []
        bug_operation_time_list= []
        bug_status_list = []
        bug_management_sn_list = []
        bug_username_operation_list = []
        url_operation_list = []
        #获取记录BUG操作过程的页面链接
        for index_bug_id, item_bug_id in enumerate(bug_id_list_total):
            url_bug_detail = "http://10.7.13.21:2000/pages/workflow/workflowChartByObject.jsf"
            querystring_bug_detail = {"workflowType":"ISU","objectId":"{bugid_value}".format(bugid_value=item_bug_id)}
            bug_detail = BeautifulSoup(get_data.get(url_bug_detail, params=querystring_bug_detail).text, "html.parser")
            url_operation_temp = bug_detail.select('iframe[id="viewFrame"]')[0].attrs['src']
            url_operation  = "http://10.7.13.21:2000" + url_operation_temp.replace("../../","/")
            url_operation_list.append(url_operation)

        #根据上一步的连接获取BUG操作记录
        for index_operation, item_operation in enumerate(url_operation_list):
            #print os.linesep
            data_operation_page = BeautifulSoup(get_data.get(item_operation).text, "html.parser")
            #print data_operation_page
            status_operation_list_temp = data_operation_page.select('table[id="procTab"] > tr > td:nth-of-type(2) > div')
            name_operation_list_temp = data_operation_page.select('table[id="procTab"] > tr > td:nth-of-type(3) > div')
            person_operation_list_temp = data_operation_page.select('table[id="procTab"] > tr > td:nth-of-type(4) > div')
            time_operation_list_temp = data_operation_page.select('table[id="procTab"] > tr > td:nth-of-type(5) > div')
            for index_operation_status, item_operation_status in enumerate(status_operation_list_temp):
                # print item_operation_status.text.strip()
                # time.sleep(1)
                if item_operation_status.text.strip() == "Verify" and name_operation_list_temp[index_operation_status].text.strip() == "驳回".decode('gbk'):
                    bug_username_operation_list.append(person_operation_list_temp[index_operation_status].text.strip())
                    bug_operation_time_list.append(time_operation_list_temp[index_operation_status].text.strip().split(" ")[0].strip())
                    bug_name_list.append(bug_name_list_total[index_operation])
                    bug_creatime_list.append(bug_creatime_list_total[index_operation])
                    bug_status_list.append(bug_status_list_total[index_operation])
                    bug_management_sn_list.append(bug_management_sn_list_total[index_operation])
                    break


        # print len(bug_name_list)
        # print len(bug_creatime_list)
        # print len(bug_status_list)
        # print len(bug_management_sn_list)
        # print len(bug_operation_time_list)
        # print len(bug_username_operation_list)

        #write to log file
        title_sheet = ['编号'.decode('gbk'), '标题'.decode('gbk'), '当前状态'.decode('gbk'), '驳回人'.decode('gbk'),'BUG创建时间'.decode('gbk'), '提交方案驳回时间'.decode('gbk')]
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
        sheet.merge_range(0, 0, 0, 5, "%s_BUG一次解决率统计".decode('gbk') % project_name_selected, formattitle)
        sheet.set_column('A:A', 17)
        sheet.set_column('B:B', 65)
        sheet.set_column('C:C', 15)
        sheet.set_column('E:F', 18)
        for index_title, item_title in enumerate(title_sheet):
            sheet.write(1, index_title, item_title, formatone)
        for index_data, item_data in enumerate(bug_management_sn_list):
            sheet.write(2 + index_data, 0, item_data, formatone)
            sheet.write(2 + index_data, 1, bug_name_list[index_data], formatone)
            sheet.write(2 + index_data, 2, bug_status_list[index_data].decode('unicode_escape'), formatone)
            sheet.write(2 + index_data, 3, bug_username_operation_list[index_data], formatone)
            sheet.write_datetime(2 + index_data, 4, datetime.datetime.strptime(bug_creatime_list[index_data], '%Y-%m-%d'), workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
            sheet.write_datetime(2 + index_data, 5, datetime.datetime.strptime(bug_operation_time_list[index_data], '%Y-%m-%d'), workbook_display.add_format({'num_format': 'yyyy-mm-dd', 'border': 1}))
        workbook_display.close()
        self.button_go.Enable()
        self.updatedisplay(("结束抓取信息，结束时间{time_end}".format(time_end=time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))).decode('gbk'))


if __name__ == '__main__':
    app = wx.App()
    frame = BugSolutionMoreThanOnce(None)
    frame.Show()
    app.MainLoop()
