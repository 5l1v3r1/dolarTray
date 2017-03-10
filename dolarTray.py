# encoding: utf-8
import win32api, win32gui
import win32con, winerror
import sys, os
from bs4 import BeautifulSoup
import requests
import urllib2
kaynak="Investing"

class MainWindow:
	def __init__(self):
		msg_TaskbarRestart = win32gui.RegisterWindowMessage("TaskbarOlustur");
		message_map = {
				msg_TaskbarRestart: self.OnRestart,
				win32con.WM_DESTROY: self.OnDestroy,
				win32con.WM_COMMAND: self.OnCommand,
				win32con.WM_USER+20 : self.OnTaskbarNotify,
		}
		# Register the Window class.
		wc = win32gui.WNDCLASS()
		hinst = wc.hInstance = win32api.GetModuleHandle(None)
		wc.lpszClassName = "dolarTray"
		wc.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW;
		wc.hCursor = win32api.LoadCursor( 0, win32con.IDC_ARROW )
		wc.hbrBackground = win32con.COLOR_WINDOW
		wc.lpfnWndProc = message_map # could also specify a wndproc.

		# Don't blow up if class already registered to make testing easier
		try:
			classAtom = win32gui.RegisterClass(wc)
		except win32gui.error, err_info:
			if err_info.winerror!=winerror.ERROR_CLASS_ALREADY_EXISTS:
				raise

		# Create the Window.
		style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
		self.hwnd = win32gui.CreateWindow( wc.lpszClassName, u"dolarTrayWin", style, \
				0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
				0, 0, hinst, None)
		win32gui.UpdateWindow(self.hwnd)
		self._DoCreateIcons()
	def _DoCreateIcons(self):
		# Try and find a custom icon
		hinst =  win32api.GetModuleHandle(None)
		#iconPathName = os.path.realpath("icon.ico")
		iconPathName = os.path.realpath("money.ico")
		#iconPathName = os.path.abspath(os.path.join( os.path.split(sys.executable)[0], "pyc.ico" ))
		#if os.path.isfile(iconPathName):
		icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
		hicon = win32gui.LoadImage(hinst, iconPathName, win32con.IMAGE_ICON, 0, 0, icon_flags)
		flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
		nid = (self.hwnd, 0, flags, win32con.WM_USER+20, hicon, u"Döviz Kurları")
		try:
			win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
			win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, \
                         (self.hwnd, 0, win32gui.NIF_INFO, win32con.WM_USER+20,\
                          hicon, "Balloon  tooltip",u"Kurları görüntülemek için çift tıklayın\nAyarlar için sağ tıklayın",200,u"Döviz Kurları"))
                        #--Kurlar güncellenip yeni bir tooltip çıkacak--
		except win32gui.error:
			# This is common when windows is starting, and this code is hit
			# before the taskbar has been created.
			print "Failed to add the taskbar icon - is explorer running?"
			# but keep running anyway - when explorer starts, we get the
			# TaskbarCreated message.

	def OnRestart(self, hwnd, msg, wparam, lparam):
		self._DoCreateIcons()

	def OnDestroy(self, hwnd, msg, wparam, lparam):
		nid = (self.hwnd, 0)
		win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
		win32gui.PostQuitMessage(0) # Terminate the app.

	def OnTaskbarNotify(self, hwnd, msg, wparam, lparam):
		if lparam==win32con.WM_LBUTTONUP:
			print u"Tık"
			
		elif lparam==win32con.WM_LBUTTONDBLCLK:
                        global kaynak
			print u"Çift Tık"
			self.balonGoster(u"Güncelleniyor...",u"Döviz Kurları (" + kaynak + ")")
                        self.kurlariAl()
                        
			
			#win32gui.DestroyWindow(self.hwnd)

			
		elif lparam==win32con.WM_RBUTTONUP:
			print u"Sağ Tık"
			menu = win32gui.CreatePopupMenu()
			win32gui.AppendMenu( menu, win32con.MF_STRING|win32con.MF_DISABLED, 0, u"Kaynak Seçimi")
			win32gui.AppendMenu( menu, win32con.MF_STRING, 1021, u"T.C.M.B.")
			win32gui.AppendMenu( menu, win32con.MF_STRING, 1022, u"Enpara")
			win32gui.AppendMenu( menu, win32con.MF_STRING, 1023, u"Investing")
			win32gui.AppendMenu( menu, win32con.MF_SEPARATOR, 0, '')
			win32gui.AppendMenu( menu, win32con.MF_STRING, 1099, u"Çıkış" )
			pos = win32gui.GetCursorPos()
			win32gui.SetForegroundWindow(self.hwnd)
			win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None)
			win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
		return 1

	def OnCommand(self, hwnd, msg, wparam, lparam):
                global kaynak
		id = win32api.LOWORD(wparam)
		if id == 1021:
			kaynak="TCMB"
			self.balonGoster(u"TCMB kurları gösterilecek",u"Döviz Kurları")
		elif id == 1022:
			kaynak="Enpara"
			self.balonGoster(u"Enpara kurları gösterilecek",u"Döviz Kurları")
		elif id == 1023:
			kaynak="Investing"
			self.balonGoster(u"Investing kurları gösterilecek",u"Döviz Kurları")
		elif id == 1099:
			print u"Kapat"
			win32gui.DestroyWindow(self.hwnd)
		else:
			print "Bilinmeyen komut -", id
                

        def balonGoster(self,mesaj,baslik,sure=150):
                nid = (self.hwnd, 0)
                hicon = win32gui.LoadImage(win32api.GetModuleHandle(None), os.path.realpath("money.ico"), win32con.IMAGE_ICON, 0, 0, win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE)
                win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, \
                         (self.hwnd, 0, win32gui.NIF_INFO, win32con.WM_USER+20,\
                          hicon, "Balloon  tooltip",mesaj,sure,baslik))
                
        def kurlariAl(self):
                global dolarAlis
                global dolarSatis
                global euroAlis
                global euroSatis
                
                if kaynak == "TCMB":
                        dolarAlis=""
                        dolarSatis=""
                elif kaynak=="Enpara":
                        dolarAlis=""
                        dolarSatis=""
                elif kaynak=="Investing":

                        opener = urllib2.build_opener()
                        opener.addheaders = [('User-agent', 'Mozilla/5.0')]
                        
                        pasteURL = "http://tr.investing.com/currencies/usd-try"

                        response = opener.open(pasteURL)
                        page = response.read().decode('utf-8')

                        parse = BeautifulSoup(page,"html.parser")
    			for dolar in parse.find_all('span', class_="pid-18-bid"):
                                dolarAlis = dolar.text.encode('utf-8')
                        for dolar in parse.find_all('span', class_="pid-18-ask"):
                                dolarSatis = dolar.text.encode('utf-8')
                                
                        pasteURL = "http://tr.investing.com/currencies/eur-try"

                        response = opener.open(pasteURL)
                        page = response.read().decode('utf-8')

                        parse = BeautifulSoup(page,"html.parser")
    			for dolar in parse.find_all('span', class_="pid-66-bid"):
                                euroAlis = dolar.text.encode('utf-8')
                        for dolar in parse.find_all('span', class_="pid-66-ask"):
                                euroSatis = dolar.text.encode('utf-8')

                self.kurlariGoster()

        def kurlariGoster(self):
                global dolarAlis
                global dolarSatis
                global euroAlis
                global euroSatis
                global kaynak
                self.balonGoster(u"Dolar\nAlış: " + dolarAlis + "      Satis: " + dolarSatis + "\nEuro\nAlis: " + euroAlis + "      Satis: " + euroSatis ,u"Döviz Kurları (" + kaynak + ")",1000)


def winmain():
	w=MainWindow()
	win32gui.PumpMessages()

if __name__=='__main__':
	winmain()
