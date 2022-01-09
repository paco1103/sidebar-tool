import os
import tkinter as tk
import wx
import json
from os import listdir, startfile
from win32com.client import Dispatch


########################################################################


class DropFile(wx.FileDropTarget):
    def __init__(self, window, buttonPanel):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.buttonPanel = buttonPanel

    def OnDropFiles(self, x, y, filenames):
        shell = Dispatch('WScript.Shell')
        for target in filenames:
            # splitext = remove extension, basename = get file name without path
            saving_path = os.path.join(
                './shortcut', os.path.splitext(os.path.basename(target))[0]+'.lnk')
            shortcut = shell.CreateShortCut(saving_path)
            # icon path is same as target file path
            shortcut.Targetpath = target
            shortcut.IconLocation = target
            shortcut.save()

        # oop, update the update panel
        self.buttonPanel.update_panel()


########################################################################


class ButtonPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, wx.ID_ANY)
        self.shortcut_path = 'shortcut'
        if not os.path.exists(self.shortcut_path):
            os.makedirs(self.shortcut_path)
        self.shortcut_end = '.lnk'
        self.item_max = 5
        self.gen_btn()

    def gen_btn(self):
        gs = wx.GridSizer(self.item_max, 1, 0, 0)
        files = listdir(self.shortcut_path)
        # sort by last modify time
        files = sorted(files,  key=lambda x: os.path.getmtime(
            os.path.join(self.shortcut_path, x)))
        shell = Dispatch("WScript.Shell")

        for i, fname in enumerate(files):
            # the max num of item show = 5
            if i > self.item_max:
                break
            if fname.endswith(self.shortcut_end):
                path_with_name = self.shortcut_path + '\\' + fname
                filePath = shell.CreateShortCut(path_with_name).Targetpath

                btn = wx.Button(self, wx.ID_ANY, label=fname.replace(
                    self.shortcut_end, ''), name=filePath)

                # TODO picture button

                btn.Bind(wx.EVT_RIGHT_UP, self.onRightClick)
                # bind the event to button
                btn.Bind(wx.EVT_BUTTON, self.onLeftClick)

                gs.Add(btn, -1, wx.EXPAND)
                self.SetSizer(gs)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self, -1, wx.EXPAND)
        self.Parent.SetSizer(sizer)
        self.Parent.Layout()
        # clear the sizer child to prevent the 'Adding a window already in a sizer, detach it first!'
        sizer.Clear()

    def clear_panel(self):
        for child in self.GetChildren():
            child.Destroy()

    def update_panel(self):
        self.clear_panel()
        self.gen_btn()

    def onLeftClick(self, event):
        # open file
        startfile(event.GetEventObject().GetName())

    def onRightClick(self, event):
        # TODO using right click menu, for remove
        #     popupmenu = wx.Menu()
        #     menuItem = wx.MenuItem(popupmenu, -1, 'Remove')
        #     popupmenu.Append(menuItem)

        #     popupmenu.Bind(wx.EVT_MENU, self.OnStuff(event, event.GetEventObject().GetName()), menuItem)
        #     #self.PopupMenu(popupmenu, event.GetPosition())

        # def OnStuff(self, event, data=None):
        #     wx.PostEvent(self, self.removeShortcut)
        self.removeShortcut(event.GetEventObject().GetLabel())

    def removeShortcut(self, shortcut_name):
        os.remove(self.shortcut_path+'/'+shortcut_name+self.shortcut_end)
        self.update_panel()

    def scale_bitmap(bitmap, width, height):
        image = wx.ImageFromBitmap(bitmap)
        image = image.Scale(width, height, wx.IMAGE_QUALITY_HIGH)
        result = wx.BitmapFromImage(image)
        return result

########################################################################


class SideBar(wx.App):
    def OnInit(self):
        # loading the setting, if first execute, create setting file
        try:
            setting = self.load_setting()
        except:
            setting = self.create_setting()

        w, h, x, y = setting['width'], setting['height'], setting['x'], setting['y']
        x = x-w
        y = y/2 - h/2

        #self.frame = wx.Frame(None, size=(1000, 1000), style=wx.BORDER_NONE)
        self.frame = wx.Frame(None, size=(w, h), style=wx.BORDER_NONE)
        # set window always on top
        self.frame.SetWindowStyle(wx.STAY_ON_TOP)

        # create the dnd function and pass the panel obj to dnd class
        self.frame.SetDropTarget(DropFile(self, ButtonPanel(self.frame)))

        self.frame.Move(wx.Point(x, y))

        # # last statment, loop until end program, infinite loop
        self.window_hide_show(tk.Tk(), x, x+w, y, y+h)
        # self.frame.Show()

        return True

    def load_setting(self):
        with open('setting.json', 'r') as f:
            return json.load(f)

    def create_setting(self):
        #default setting, one monitor with 1920*1080
        setting = {
            "width": 50,
            "height": 350,
            "x": 1920,
            "y": 1080
        }
        with open('setting.json', 'w') as f:
            json.dump(setting, f, ensure_ascii=False, indent=4)
        return setting

    # control the window hide and show, infinite loop
    def window_hide_show(self, tk, start_pos_x, end_pos_x, start_pos_y, end_pos_y):
        position = tk.winfo_pointerxy()  # (x, y)
        if position[0] > start_pos_x and position[0] < end_pos_x and position[1] > start_pos_y and position[1] < end_pos_y:
            self.frame.SetWindowStyle(wx.STAY_ON_TOP)
            self.frame.Show()
        else:
            self.frame.Hide()
        wx.CallLater(500, lambda: self.window_hide_show(
            tk, start_pos_x, end_pos_x, start_pos_y, end_pos_y))


sideBar = SideBar()
sideBar.MainLoop()
