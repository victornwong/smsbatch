#!/usr/bin/python
# -*- coding: utf-8 -*-
'''
SMS Batch Processor
Version 0.1
Written by Victor Wong
Since 03/07/2015

Notes:

'''
import wx,os,xlrd
import wx.lib.mixins.listctrl as listmix
import sqlite3 as lite

class EditableListCtrl(wx.ListCtrl, listmix.TextEditMixin):
	''' TextEditMixin allows any column to be edited. '''
	def __init__(self, parent, ID=wx.ID_ANY, pos=wx.DefaultPosition,size=(-1,350), style=0):
		wx.ListCtrl.__init__(self, parent, ID, pos, size, style)
		listmix.TextEditMixin.__init__(self)

class MainWindow(wx.Frame):
	def __init__(self, *args, **kwargs):
		super(MainWindow, self).__init__(*args, **kwargs) 
		self.InitUI()

	def InitUI(self):
		panel = wx.Panel(self,wx.ID_ANY)
		panel.SetBackgroundColour((0x25,0x43,0x6A))
		menubar = wx.MenuBar()
		fileMenu = wx.Menu()
		loadvw_itm = fileMenu.Append(wx.ID_ANY, "List previous")
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
		self.SetMenuBar(menubar)

		#self.list_ctrl = wx.ListCtrl(panel, size=(-1,300), style=wx.LC_REPORT|wx.BORDER_SUNKEN)
		self.list_ctrl = EditableListCtrl(panel, style=wx.LC_REPORT)
		self.list_ctrl.InsertColumn(0,"REC",width=40)
		self.list_ctrl.InsertColumn(1,"Voucher",width=50)
		self.list_ctrl.InsertColumn(2,"Customer",width=180)
		self.list_ctrl.InsertColumn(3,"Phone",width=80)
		self.list_ctrl.InsertColumn(4,"Message")
		self.list_ctrl.InsertColumn(5,"Sent",width=70)
		self.list_ctrl.InsertColumn(6,"Resend",width=70)

		#st1 = wx.StaticText(panel,label="Worksheet",style=wx.ALIGN_LEFT)

		self.sendsms_btn = wx.Button(panel,label="Start send SMS")
		self.sendsms_btn.SetBackgroundColour((0xE9,0x19,0x49))
		self.sendsms_btn.SetForegroundColour(wx.WHITE)
		self.sendsms_btn.Bind(wx.EVT_BUTTON,self.StartSendSMS)

		btn2 = wx.Button(panel,label="Clear worksheet")
		btn2.SetBackgroundColour((0xE9,0x19,0x49))
		btn2.SetForegroundColour(wx.WHITE)
		btn2.Bind(wx.EVT_BUTTON,self.ClearWorksheet)

		btn3 = wx.Button(panel,label="Upload worksheet")
		btn3.SetBackgroundColour((0xE9,0x19,0x49))
		btn3.SetForegroundColour(wx.WHITE)
		btn3.Bind(wx.EVT_BUTTON,self.OnUploadworksheet)

		vbox = wx.BoxSizer(wx.VERTICAL)

		hbox = wx.BoxSizer(wx.HORIZONTAL)
		hbox.Add(btn3,0,wx.ALL,5)
		hbox.Add(self.sendsms_btn,0,wx.ALL,5)
		hbox.Add(btn2,0,wx.ALL,5)

		vbox.Add(hbox,0,wx.ALL|wx.EXPAND,2)
		vbox.Add(self.list_ctrl,0,wx.ALL|wx.EXPAND,2)

		panel.SetSizer(vbox)

		self.SetSize((800, 420))
		self.SetTitle('SMS Batch Processor')
		self.Centre()
		self.Show(True)

	def StartSendSMS(self,e):
		mainlist = []
		count = self.list_ctrl.GetItemCount()
		for row in range(count):
			wop = []
			for col in range(1,6):
				itm = self.list_ctrl.GetItem(row,col)
				ival = itm.GetText()
				wop.append(ival)

			mainlist.append(wop)

		sqlstm = ""

		for ki in mainlist:
			sqlstm += "insert into smsr (voucherno,customer) values ('" + ki[0] + "','" + ki[1] + "');"

		con = lite.connect("records.db")
		cur = con.cursor()
		cur.executescript(sqlstm)
		con.commit()
		con.close()

	def ClearWorksheet(self,e):
		self.list_ctrl.DeleteAllItems()

	def ListRecords(self,e):
		# Load from sqlite the previous records
		self.list_ctrl.DeleteAllItems()
		self.sendsms_btn.Disable()

		con = None

		try:
			con = lite.connect("records.db")
			cur = con.cursor()
			try:
				cur.execute("CREATE TABLE smsr(voucherno TEXT, customer TEXT, phone TEXT, message TEXT, sent TEXT, resend TEXT);")
			except lite.Error,e:
				print "Err %s:" % e.args[0]

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
			self.sendsms_btn.Enable()
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
			'''
			,
			"Worksheet template info",wx.OK | wx.ICON_INFORMATION)

	def AboutBox(self,e):
		wx.MessageBox("SMS Batch Processor v0.1\nWritten by Victor Wong\nSince 03/07/2015",
			"About",wx.OK | wx.ICON_INFORMATION)

	def OnQuit(self, e):
		self.Close()

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
