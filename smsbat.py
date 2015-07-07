#!/usr/bin/python
# -*- coding: utf-8 -*-
'''
SMS Batch Processor
Version 0.1
Written by Victor Wong
Since 03/07/2015

Notes:
Uses OneWaySMS API

'''
import wx,os,xlrd
import wx.lib.mixins.listctrl as listmix
import sqlite3 as lite

DBNAME = "records.db"

lheader = ["REC","Voucher","Customer","Phone","Message","Sent","Resend","TS"]
lhwidth = [70,70,180,80,180,70,70,30]

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
		self.InitUI()

	def checkDatabase(self):
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
					sent VARCHAR(30), resend VARCHAR(30), nosend TINYINT);
					""")
			except lite.Error,e:
				print "Err %s:" % e.args[0]

		except lite.Error,e:
			print "Err %s:" % e.args[0]
			wx.MessageBox("ERR: Cannot read database","ERROR",wx.OK | wx.ICON_ERROR)

		finally:
			if con:
				con.close()

	def InitUI(self):
		panel = wx.Panel(self,wx.ID_ANY)
		panel.SetBackgroundColour((0x25,0x43,0x6A))
		menubar = wx.MenuBar()
		fileMenu = wx.Menu()
		loadvw_itm = fileMenu.Append(wx.ID_ANY, "List previous")
		clrdb_itm = fileMenu.Append(wx.ID_ANY, "Clear database")
		quit_itm = fileMenu.Append(wx.ID_EXIT, 'Quit', 'Quit application')
		menubar.Append(fileMenu, "&Function")

		helpMenu = wx.Menu()
		hlp_itm = helpMenu.Append(wx.ID_ANY,"Worksheet template")
		abt_itm = helpMenu.Append(wx.ID_ANY,"About")
		menubar.Append(helpMenu,"&Help")

		self.Bind(wx.EVT_MENU, self.OnQuit, quit_itm)
		self.Bind(wx.EVT_MENU, self.AboutBox, abt_itm)
		self.Bind(wx.EVT_MENU, self.TemplateHelpBox, hlp_itm)
		self.Bind(wx.EVT_MENU, self.ListRecords, loadvw_itm)
		self.Bind(wx.EVT_MENU, self.ClearDatabase, clrdb_itm)
		self.SetMenuBar(menubar)

		#self.list_ctrl = wx.ListCtrl(panel, size=(-1,300), style=wx.LC_REPORT|wx.BORDER_SUNKEN)
		self.list_ctrl = EditableListCtrl(panel, style=wx.LC_REPORT|wx.BORDER_SUNKEN)

		for i in range(len(lheader)):
			self.list_ctrl.InsertColumn(i,lheader[i],width=lhwidth[i])

		#st1 = wx.StaticText(panel,label="Worksheet",style=wx.ALIGN_LEFT)

		btnid = ["sendallsms","resendsms","clearworksheet","uploadworksheet","deleteentry"]
		btnlabel = ["Send all SMS", "Resend SMS", "Clear worksheet", "Upload worksheet", "Delete entry"]
		btnfunc = [self.StartSendSMS, self.ResendSMS, self.ClearWorksheet, self.OnUploadworksheet, self.DeleteEntry]
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
		hbox.Add(self.btns["uploadworksheet"],0,wx.ALL,5)
		hbox.Add(self.btns["sendallsms"],0,wx.ALL,5)
		hbox.Add(self.btns["resendsms"],0,wx.ALL,5)
		
		hbox2 = wx.BoxSizer(wx.HORIZONTAL)
		hbox2.Add(self.btns["deleteentry"],0,wx.ALL,5)
		hbox2.Add(self.btns["clearworksheet"],0,wx.ALL,5)

		vbox.Add(hbox,0,wx.ALL|wx.EXPAND,2)
		vbox.Add(self.list_ctrl,0,wx.ALL|wx.EXPAND,2)
		vbox.Add(hbox2,0,wx.ALL|wx.EXPAND,2)

		panel.SetSizer(vbox)

		self.SetSize((800, 440))
		self.SetTitle('SMS Batch Processor')
		self.Centre()
		self.Show(True)

	def UpdateListToDatabase(self,iwhat):
		mainlist = []
		count = self.list_ctrl.GetItemCount()
		for row in range(count):
			wop = []
			for col in range(0,len(lheader)):
				itm = self.list_ctrl.GetItem(row,col)
				ival = itm.GetText()
				wop.append(ival)

			mainlist.append(wop)

		sqlstm = "begin;"

		#"Voucher","Customer","Phone","Message","Sent","Resend","TS"]
		for ki in mainlist:
			if iwhat == True:
				sqlstm += "insert into smsr (origid,voucherno,customer,phone,message,sent,resend,nosend) values (NULL,'" + ki[1] + "','" + ki[2] + "','" + ki[3] + "','" + ki[4] + "','" + ki[5] + "','" + ki[6] + "',0);"
			else:
				sqlstm += "update smsr set voucherno='" + ki[1] + "', customer='" + ki[2] + "', phone='" + ki[3] + "', message='" + ki[4] + "', sent='" + ki[5] + "', resend='" + ki[6] + "' where origid=" + ki[0] + ";"

		sqlstm += "end;"
		dbExecuter(sqlstm)

	def DeleteEntry(self,e):
		kk = get_selected_items(self.list_ctrl)

		if len(kk) == 0:
			return

		dlg = wx.MessageDialog(None, 'Are you sure you want remove the selected?', 'Question',wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
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
		dlg = wx.MessageDialog(None, 'Are you sure you want clear the database?', 'Question',wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
		res = dlg.ShowModal()

		if res == wx.ID_YES:
			dbExecuter("delete from smsr")
			print "Database cleared"
			self.list_ctrl.DeleteAllItems() # once db clear, delete all listbox entries

	def StartSendSMS(self,e):
		# loop list and send sms
		#self.UpdateListToDatabase(self.newupload) # save them rows into db
		print "send all sms"

	def ResendSMS(self,e):
		print "resend sms"

	def ClearWorksheet(self,e):
		self.list_ctrl.DeleteAllItems()

	def ListRecords(self,e):
		#self.sendsms_btn.Disable()
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

			flds = ["voucherno","customer","phone","message","sent","resend","nosend"]

			for d in drws:
				self.list_ctrl.InsertStringItem(index,str(d["origid"]))

				for i in range(len(flds)):
					self.list_ctrl.SetStringItem(index,i+1,str(d[flds[i]]))

				if index % 2:
					self.list_ctrl.SetItemBackgroundColour(index,"white")
				else:
					self.list_ctrl.SetItemBackgroundColour(index,(0x2A,0x9D,0xD5))

				index += 1;

		except lite.Error,e:
			print "Err %s:" % e.args[0]
			wx.MessageBox("ERR: Cannot read database","ERROR",wx.OK | wx.ICON_ERROR)

		finally:
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
			#self.sendsms_btn.Enable()
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
							self.list_ctrl.InsertStringItem(index,str(index))
							self.list_ctrl.SetStringItem(index,2,str(cellval))
						else:
							self.list_ctrl.SetStringItem(index,curr_cell+2,str(cellval))

						if index % 2:
							self.list_ctrl.SetItemBackgroundColour(index,"white")
						else:
							self.list_ctrl.SetItemBackgroundColour(index,(0x2A,0x9D,0xD5))

					index += 1;

		except xlrd.XLRDError:
			wx.MessageBox("ERR: Cannot process file","ERROR",wx.OK | wx.ICON_ERROR)

	def TemplateHelpBox(self,e):
		wx.MessageBox(
			'''
Worksheet template must be in MS-Excel XP/2003 (xls) format ONLY.

Internal record numbers are auto-increment.
			'''
			,
			"Worksheet template info",wx.OK | wx.ICON_INFORMATION)

	def AboutBox(self,e):
		wx.MessageBox("SMS Batch Processor v0.1\nWritten by Victor Wong\nSince 03/07/2015",
			"About",wx.OK | wx.ICON_INFORMATION)

	def OnQuit(self, e):
		self.Close()

def zebra_paint(list_control):
	count = list_control.GetItemCount()
	for row in range(count):
		if row % 2:
			list_control.SetItemBackgroundColour(row,"white")
		else:
			list_control.SetItemBackgroundColour(row,(0x2A,0x9D,0xD5))

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

'''
class UploadWorksheetDialog(wx.Dialog):
	def __init__(self, *args, **kw):
		super(UploadWorksheetDialog, self).__init__(*args, **kw) 
		self.InitUI()
		self.SetSize((250, 200))
		self.SetTitle("Upload worksheet")

	def InitUI(self):
		pnl = wx.Panel(self)
		vbox = wx.BoxSizer(wx.VERTICAL)
		upload_b = wx.Button(self,label="Select upload")
		self.Bind(wx.EVT_BUTTON,self.OnWorkUpload,upload_b)
		vbox.Add(upload_b)

	def OnWorkUpload(self,e):
		e.Veto()

'''
