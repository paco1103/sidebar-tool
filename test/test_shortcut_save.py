import os, winshell, win32com.client
desktop = winshell.desktop()
#desktop = r"path to where you wanna put your .lnk file"
path = os.path.join(desktop, 'File Shortcut Demo.lnk')
path = saving_path = os.path.join('./shortcut', os.path.splitext(target)[0]+'.lnk')
target = r"C:\Users\user\Desktop\123.txt" 
icon = r"C:\Users\user\Desktop\123.txt"
shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(path)
shortcut.Targetpath = target
shortcut.IconLocation = icon
#shortcut.save()

print(desktop)
print(path)
print(shortcut)