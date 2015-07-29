#!/usr/bin/python
# -*- coding: utf-8 -*-
'''
SMS Batch Sender
Version 0.1
Written by Victor Wong
Since 03/07/2015

Notes:
Uses onewaysms API Version 1.2 12/03/2013

'''
import wx,os,xlrd
import httplib, urllib, ConfigParser
import wx.lib.mixins.listctrl as listmix
import sqlite3 as lite

DBNAME = "records.db"
PROGRAM_TITLE = "SMS Batch Sender"
PROGRAM_VERSION = "v0.1"
CONFIG_FILENAME = "config.ini"
CONFIG_SECTION = "onewaysms"

OR_COL = 0
VO_COL = 1
CS_COL = 2
PH_COL = 3
MS_COL = 4
ST_COL = 5
RS_COL = 6
TS_COL = 7
RESP_COL = 8

lheader = ["REC","Voucher","Customer","Phone","Message","Sent","Resend","TS","Resp"]
lhwidth = [70,70,180,80,180,70,70,30,50]

mconfig = ConfigParser.SafeConfigParser()

#gateway_username = gateway_password = ""

class SMSGatewaySettingDialog(wx.Dialog):
	def __init__(self, *args, **kw):
		super(SMSGatewaySettingDialog, self).__init__(*args, **kw)
		self.InitUI()
		self.SetTitle("SMS gateway settings")

	def InitUI(self):
		pnl = wx.Panel(self)

		self.gwurl = wx.TextCtrl(pnl, size=(140,-1))
		self.gwuname = wx.TextCtrl(pnl, size=(140,-1))
		self.gwpaswd = wx.TextCtrl(pnl, size=(140,-1))
		self.gwport = wx.TextCtrl(pnl, size=(140,-1))

		try:
			self.gwurl.SetValue(mconfig.get(CONFIG_SECTION,"url"))
			self.gwuname.SetValue(mconfig.get(CONFIG_SECTION,"username"))
			self.gwpaswd.SetValue(mconfig.get(CONFIG_SECTION,"password"))
			self.gwport.SetValue(mconfig.get(CONFIG_SECTION,"port"))
		except Exception, e:
			pass

		button = wx.Button(pnl, label="Save")
		button.Bind(wx.EVT_BUTTON, self.SaveGWSetting)

		gsizer = wx.GridBagSizer(4,2)
		gsizer.Add(wx.StaticText(pnl,label="GW URL"),(0,0))
		gsizer.Add(self.gwurl,(0,1))

		gsizer.Add(wx.StaticText(pnl,label="GW port"),(1,0))
		gsizer.Add(self.gwport,(1,1))

		gsizer.Add(wx.StaticText(pnl,label="GW username"),(2,0))
		gsizer.Add(self.gwuname,(2,1))

		gsizer.Add(wx.StaticText(pnl,label="GW password"),(3,0))
		gsizer.Add(self.gwpaswd,(3,1))

		gsizer.Add(button,(4,0),(4,2),flag=wx.EXPAND)
		
		border = wx.BoxSizer()
		border.Add(gsizer,1,wx.ALL | wx.EXPAND, 10)

		pnl.SetSizer(border)
		pnl.SetAutoLayout(True) # auto-resize panel->dialog
		border.Fit(pnl)
		pnl.Fit()
		self.Fit()

	def SaveGWSetting(self,e):
		fo = open(CONFIG_FILENAME,"w")
		mconfig.set(CONFIG_SECTION, "url", self.gwurl.GetValue().strip())
		mconfig.set(CONFIG_SECTION, "username", self.gwuname.GetValue().strip())
		mconfig.set(CONFIG_SECTION, "password", self.gwpaswd.GetValue().strip())
		mconfig.set(CONFIG_SECTION, "port", self.gwport.GetValue().strip())
		mconfig.write(fo)
		fo.close()
		self.Destroy()

class EditableListCtrl(wx.ListCtrl, listmix.TextEditMixin):
	''' TextEditMixin allows any column to be edited. '''
	def __init__(self, parent, ID=wx.ID_ANY, pos=wx.DefaultPosition,size=(-1,350), style=0):
		wx.ListCtrl.__init__(self, parent, ID, pos, size, style)
		listmix.TextEditMixin.__init__(self)
		self.Bind(wx.EVT_LIST_BEGIN_LABEL_EDIT, self.OnBeginLabelEdit)

	def OnBeginLabelEdit(self, event):
		if event.m_col == 0: # record-no cannot edit
			event.Veto()
		else:
			event.Skip()

class MainWindow(wx.Frame):

	def __init__(self, *args, **kwargs):
		super(MainWindow, self).__init__(*args, **kwargs)
		self.newupload = False
		self.checkDatabase()
		self.loadConfig()
		self.InitUI()

	def checkDatabase(self):
		''' Check database and create tables if not exist '''
		con = None
		try:
			con = lite.connect(DBNAME)
			cur = con.cursor()
			try:
				cur.execute(
					"""
					CREATE TABLE smsr(origid INTEGER PRIMARY KEY AUTOINCREMENT,
					voucherno VARCHAR(50), customer VARCHAR(300),
					phone VARCHAR(30), message VARCHAR(320),
					sent VARCHAR(30), resend VARCHAR(30), nosend TINYINT,
					gwresponse VARCHAR(255)
					);
					""")
			except lite.Error,e:
				#print "Err %s:" % e.args[0]
				pass

		except lite.Error,e:
			print "Err %s:" % e.args[0]
			wx.MessageBox("ERR: Cannot read database","ERROR",wx.OK | wx.ICON_ERROR)

		finally:
			if con:
				con.close()

	def loadConfig(self):
		mconfig.read(CONFIG_FILENAME)
		try:
			kk = mconfig.get(CONFIG_SECTION,"username")
		except ConfigParser.NoSectionError, e:
			mconfig.add_section(CONFIG_SECTION) # if no config found, add section for later

	def InitUI(self):
		panel = wx.Panel(self,wx.ID_ANY)
		panel.SetBackgroundColour((0x25,0x43,0x6A))
		menubar = wx.MenuBar()
		fileMenu = wx.Menu()
		loadvw_itm = fileMenu.Append(wx.ID_ANY, "List previous")
		setting_itm = fileMenu.Append(wx.ID_ANY,"Gateway setting")
		chkbalance_itm = fileMenu.Append(wx.ID_ANY,"Check credit balance")
		fileMenu.AppendSeparator()
		quit_itm = fileMenu.Append(wx.ID_EXIT, 'Quit', 'Quit application')
		menubar.Append(fileMenu, "&Function")

		helpMenu = wx.Menu()
		hlp_itm = helpMenu.Append(wx.ID_ANY,"Worksheet template")
		abt_itm = helpMenu.Append(wx.ID_ANY,"About")
		helpMenu.AppendSeparator()
		clrdb_itm = helpMenu.Append(wx.ID_ANY, "Clear database")
		menubar.Append(helpMenu,"&Help")

		self.Bind(wx.EVT_MENU, self.OnQuit, quit_itm)
		self.Bind(wx.EVT_MENU, self.AboutBox, abt_itm)
		self.Bind(wx.EVT_MENU, self.TemplateHelpBox, hlp_itm)
		self.Bind(wx.EVT_MENU, self.ListRecords, loadvw_itm)
		self.Bind(wx.EVT_MENU, self.ClearDatabase, clrdb_itm)
		self.Bind(wx.EVT_MENU, self.Mn_GatewaySetting, setting_itm)
		self.Bind(wx.EVT_MENU, self.CheckCreditBalance, chkbalance_itm)
		self.SetMenuBar(menubar)

		#self.list_ctrl = wx.ListCtrl(panel, size=(-1,300), style=wx.LC_REPORT|wx.BORDER_SUNKEN)
		self.list_ctrl = EditableListCtrl(panel, style=wx.LC_REPORT|wx.BORDER_SUNKEN)

		for i in range(len(lheader)):
			self.list_ctrl.InsertColumn(i,lheader[i],width=lhwidth[i])

		#st1 = wx.StaticText(panel,label="Worksheet",style=wx.ALIGN_LEFT)

		btnid = ["sendallsms","resendsms","clearworksheet","uploadworksheet","deleteentry","saveworksheet","newentry"]
		btnlabel = ["Send SMS", "Resend SMS", "Clear worksheet", "Upload worksheet", "Delete entry","Save worksheet","New entry"]
		btnfunc = [self.StartSendSMS, self.ResendSMS, self.ClearWorksheet, self.OnUploadworksheet, self.DeleteEntry, self.SaveWorksheet, self.NewEntry]
		self.btns = {}

		for i in range(len(btnid)):
			btn = wx.Button(panel, label=btnlabel[i])
			btn.SetBackgroundColour((0xE9,0x19,0x49))
			btn.SetForegroundColour(wx.WHITE)

			if btnfunc[i] != None:
				btn.Bind(wx.EVT_BUTTON,btnfunc[i])

			self.btns[btnid[i]] = btn

		vbox = wx.BoxSizer(wx.VERTICAL)

		hbox = wx.BoxSizer(wx.HORIZONTAL)
		hbox.Add(self.btns["uploadworksheet"],0,wx.ALL,3)
		hbox.Add(self.btns["sendallsms"],0,wx.ALL,3)
		hbox.Add(self.btns["resendsms"],0,wx.ALL,3)
		self.btns["resendsms"].Show(False)
		
		hbox2 = wx.BoxSizer(wx.HORIZONTAL)
		hbox2.Add(self.btns["saveworksheet"],0,wx.ALL,3)
		hbox2.Add(self.btns["newentry"],0,wx.ALL,3)
		hbox2.Add(self.btns["deleteentry"],0,wx.ALL,3)
		hbox2.Add(self.btns["clearworksheet"],0,wx.ALL,3)

		#self.logbox = wx.TextCtrl(panel, size=(-1, 150), style = wx.TE_MULTILINE|wx.TE_READONLY|wx.TE_AUTO_URL)
		self.logbox = wx.TextCtrl(panel, size=(-1, -1), style = wx.TE_MULTILINE|wx.TE_READONLY|wx.TE_AUTO_URL)

		vbox.Add(hbox,0,wx.ALL|wx.EXPAND,2)
		vbox.Add(self.list_ctrl,0,wx.ALL|wx.EXPAND,2)
		vbox.Add(hbox2,0,wx.ALL|wx.EXPAND,2)
		#vbox.Add(self.logbox,1,wx.ALL|wx.EXPAND,2)
		vbox.Add(self.logbox,1,wx.EXPAND)

		panel.SetSizerAndFit(vbox)

		self.SetSize((800, 600))
		self.SetTitle(PROGRAM_TITLE + " " + PROGRAM_VERSION)
		self.Centre()
		self.Show(True)

	def Mn_GatewaySetting(self,e):
		sdlg = SMSGatewaySettingDialog(None,title="something")
		sdlg.ShowModal()
		sdlg.Destroy()
		self.loadConfig() # reload config.ini for changes made

	def CheckCreditBalance(self,e):
		try:
			unm = mconfig.get(CONFIG_SECTION,"username")
			pws = mconfig.get(CONFIG_SECTION,"password")
			gurl = mconfig.get(CONFIG_SECTION,"url")
			gport = int(mconfig.get(CONFIG_SECTION,"port"))
		except Exception, e:
			wx.MessageBox("ERR: invalid gateway configuration","ERROR", wx.OK | wx.ICON_ERROR)
			return

		#r = urlopen("http://gateway.onewaysms.com.my:10001/bulkcredit.aspx?apiusername=&apipassword=")
		#httplib.HTTPConnection.debuglevel = 1
		conn = httplib.HTTPConnection(gurl,gport) # hardcoded port 10001
		chkcredit_url = "/bulkcredit.aspx?apiusername=" + unm + "&apipassword=" + pws
		conn.request("GET",chkcredit_url)
		response = conn.getresponse()
		rdata = response.read()
		#self.logbox.AppendText(str(response.status) + " " + response.reason)
		self.logbox.AppendText("\nCredit left: " + rdata)

	def UpdateListToDatabase(self,iwhat):
		mainlist = []
		count = self.list_ctrl.GetItemCount()
		for row in range(count):
			wop = []
			for col in range(0,len(lheader)+1):
				itm = self.list_ctrl.GetItem(row,col)
				ival = itm.GetText()
				wop.append(ival)

			mainlist.append(wop)

		sqlstm = "begin;"

		for ki in mainlist:
			if ki[0] == "0": # if rec no. is 0, do new insert into db
				sqlstm += "insert into smsr (origid,voucherno,customer,phone,message,sent,resend,nosend) values (NULL,'" + ki[1] + "','" + ki[2] + "','" + ki[3] + "','" + ki[4] + "','" + ki[5] + "','" + ki[6] + "',0);"
			else: # just update by rec no.
				sqlstm += "update smsr set voucherno='" + ki[1] + "', customer='" + ki[2] + "', phone='" + ki[3] + "', message='" + ki[4] + "', sent='" + ki[5] + "', resend='" + ki[6] + "',nosend=" + ki[7] + " where origid=" + ki[0] + ";"

		sqlstm += "end;"
		dbExecuter(sqlstm)

		self.ListRecords(self) # reload worksheet when recs are inserted into db

	def SaveWorksheet(self,e):
		self.UpdateListToDatabase(self.newupload)
		wx.MessageBox("Entries saved..","Info",wx.OK | wx.ICON_INFORMATION)

	def NewEntry(self,e):
		self.list_ctrl.InsertStringItem(0,"0")

	def DeleteEntry(self,e):
		kk = get_selected_items(self.list_ctrl)

		if len(kk) == 0:
			return

		dlg = wx.MessageDialog(None, "Are you sure you want remove the selected?", "Question",wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
		res = dlg.ShowModal()

		if res == wx.ID_YES:
			dosql = False
			sqlstm = "begin;"

			if self.newupload == False:
				dosql = True
				for i in range(len(kk)):
					itm = self.list_ctrl.GetItem(kk[i],0) # get origid
					sqlstm += "delete from smsr where origid=" + itm.GetText() + ";"

			kk.sort(reverse=True) # reverse selected items list, so it'll delete properly descending
			for i in range(len(kk)):
				self.list_ctrl.DeleteItem(kk[i])

			zebra_paint(self.list_ctrl)

			if dosql:
				sqlstm += "end;"
				dbExecuter(sqlstm)

	def ClearDatabase(self,e):
		dlg = wx.MessageDialog(None, "Are you sure you want clear the database?", "Question",wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
		res = dlg.ShowModal()

		if res == wx.ID_YES:
			dbExecuter("delete from smsr; delete from sqlite_sequence where name='smsr';")
			self.list_ctrl.DeleteAllItems() # once db clear, delete all listbox entries
			wx.MessageBox("Database cleared..","Info",wx.OK | wx.ICON_INFORMATION)

	def StartSendSMS(self,e):
		try:
			unm = mconfig.get(CONFIG_SECTION,"username")
			pws = mconfig.get(CONFIG_SECTION,"password")
			gurl = mconfig.get(CONFIG_SECTION,"url")
			gport = int(mconfig.get(CONFIG_SECTION,"port"))
		except Exception, e:
			wx.MessageBox("ERR: invalid gateway configuration","ERROR", wx.OK | wx.ICON_ERROR)
			return

		dlg = wx.MessageDialog(None, "Going to send SMS to selected entries?", "Question", wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
		res = dlg.ShowModal()

		if res == wx.ID_YES:
			kk = get_selected_items(self.list_ctrl)
			for i in range(len(kk)): # loop selected items and send sms
				sd = []
				for b in range(0,9):
					itm = self.list_ctrl.GetItem(kk[i],b)
					ival = itm.GetText()
					sd.append(ival)

				if sd[PH_COL].strip() is not u"":
					conn = httplib.HTTPConnection(gurl,gport)
					sendsms_url = "/api.aspx?apiusername=xxx" + unm + "&apipassword=" + pws + "&mobileno=" + sd[PH_COL].strip() + "&senderid=INFO&languagetype=1&message" + urllib.urlencode({"":sd[MS_COL].strip()})

					conn.request("GET",sendsms_url)
					response = conn.getresponse()
					rdata = response.read()
					#self.logbox.AppendText("\n" + str(response.status) + " " + response.reason)
					#self.logbox.AppendText("\n" + rdata)

					response_str = "";
					sendok = False

					if rdata == "-100":
						response_str = "CONFIG ERROR"
					elif rdata == "-200":
						response_str = "SENDERID ERROR"
					elif rdata == "-300":
						response_str = "INVALID MOBILE"
					elif rdata == "-400":
						response_str = "INVALID LANGUAGE"
					elif rdata == "-500":
						response_str = "INVALID MSG"
					elif rdata == "-600":
						response_str = "INSUFFICIENT CREDIT"
					else:
						response_str = rdata # save the MT ID
						sendok = True

					its = 0
					if sendok:
						its = int(sd[TS_COL]) + 1
						self.list_ctrl.SetStringItem(kk[i],TS_COL,str(its))

					self.list_ctrl.SetStringItem(kk[i],RESP_COL,str(response_str))
					#self.logbox.AppendText("\n" + sendsms_url)
					self.logbox.AppendText("\nSend " + sd[CS_COL] + "(" + sd[PH_COL] + ") " + sd[VO_COL] + " : " + str(its) + " count.\n" + "Response: " + response_str)

			self.UpdateListToDatabase(self.newupload) # update entries to db

	def ResendSMS(self,e):
		pass

	def ClearWorksheet(self,e):
		self.list_ctrl.DeleteAllItems()

	def ListRecords(self,e):
		# Load from sqlite the previous records
		self.list_ctrl.DeleteAllItems()
		self.newupload = False # list prev recs - will use update instead of insert later
		con = None

		try:
			con = lite.connect(DBNAME)
			con.row_factory = lite.Row
			cur = con.cursor()
			cur.execute("select * from smsr;")
			drws = cur.fetchall()
			index = 0

			flds = ["voucherno","customer","phone","message","sent","resend","nosend","gwresponse"]

			for d in drws:
				self.list_ctrl.InsertStringItem(index,str(d["origid"]))

				for i in range(len(flds)):
					lks = str(d[flds[i]])
					if d[flds[i]] == None:
						lks = ""

					self.list_ctrl.SetStringItem(index,i+1,lks)

				index += 1;

		except lite.Error,e:
			print "Err %s:" % e.args[0]
			wx.MessageBox("ERR: Cannot read database","ERROR",wx.OK | wx.ICON_ERROR)

		finally:

			zebra_paint(self.list_ctrl)
			if con:
				con.close()

	def OnUploadworksheet(self,e):
		#ku = UploadWorksheetDialog(None,title="") ku.ShowModal() ku.Destroy()
		wildcard = "MS Excel (*.xls)|*.xls|All files (*.*)|*.*"
		udlg = wx.FileDialog(None,"Choose worksheet",os.getcwd(),"",wildcard,wx.OPEN)

		if udlg.ShowModal() == wx.ID_OK:
			self.ProcessWorksheet(udlg.GetPath())

		udlg.Destroy()

	def ProcessWorksheet(self,ifilename):
		try:
			wkb = xlrd.open_workbook(ifilename)
			sheets = wkb.sheet_names()
			index = 0
			self.newupload = True # will insert records instead of update
			self.list_ctrl.DeleteAllItems() # delete all list items before inserting

			# go through every worksheet in the workbook. import 'em rows according to template format
			for wkn in sheets:
				wks = wkb.sheet_by_name(wkn)
				nrows = wks.nrows - 1
				ncells = wks.ncols - 1
				curr_row = 0 # start from row 2, row 1 are the headers
				while curr_row < nrows:
					curr_row += 1
					row = wks.row(curr_row)
					curr_cell = -1
					while curr_cell < ncells:
						curr_cell += 1
						celltype = wks.cell_type(curr_row,curr_cell)
						cellval = wks.cell_value(curr_row,curr_cell)

						if curr_cell == 0:
							self.list_ctrl.InsertStringItem(index,"0") # rec no. always 0 for uploaded worksheet - to be used in UpdateListToDatabase for insert
							self.list_ctrl.SetStringItem(index,2,str(cellval))
						else:
							self.list_ctrl.SetStringItem(index,curr_cell+2,str(cellval))

					index += 1;

		except xlrd.XLRDError:
			wx.MessageBox("ERR: Cannot process file","ERROR",wx.OK | wx.ICON_ERROR)

		finally:
			zebra_paint(self.list_ctrl)

	def TemplateHelpBox(self,e):
		wx.MessageBox(
			'''
Worksheet template must be in MS-Excel XP/2003 (xls) format ONLY.
			'''
			,
			"Worksheet template info",wx.OK | wx.ICON_INFORMATION)

	def AboutBox(self,e):
		wx.MessageBox(PROGRAM_TITLE + " " + PROGRAM_VERSION + "\nWritten by Victor Wong\nSince 03/07/2015",
			"About",wx.OK | wx.ICON_INFORMATION)

	def OnQuit(self, e):
		self.Close()

def zebra_paint(list_control):
	count = list_control.GetItemCount()
	for row in range(count):
		if row % 2:
			list_control.SetItemBackgroundColour(row,"white")
		else:
			list_control.SetItemBackgroundColour(row,(0x4D,0xCE,0xA9))

def get_selected_items(list_control):
	selection = []
	# start at -1 to get the first selected item
	current = -1
	while True:
		next = GetNextSelected(list_control, current)
		if next == -1:
			return selection

		selection.append(next)
		current = next

def GetNextSelected(list_control, current):
	"""Returns next selected item, or -1 when no more"""
	return list_control.GetNextItem(current,wx.LIST_NEXT_ALL,wx.LIST_STATE_SELECTED)

def dbExecuter(tsqlstm):
	con = lite.connect(DBNAME)
	cur = con.cursor()
	cur.executescript(tsqlstm)
	con.commit()
	con.close()

def main():
	ex = wx.App()
	MainWindow(None)
	ex.MainLoop()    

if __name__ == '__main__':
	main()
