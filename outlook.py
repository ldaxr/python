import os
import time
import win32com.client as win32
from PIL import ImageGrab

#pip install pillow
#pip install PyEmail
#pip install exchangelib

def GrabURL(url, filename, x1, y1, x2, y2) :
    os.system('"rundll32" url.dll,FileProtocolHandler ' + url)
    time.sleep(5)
    im = ImageGrab.grab((x1, y1, x2, y2)) 
    im.save(filename,'JPEG') 
    
GrabURL('http://www.baidu.com', 'file1.jpg', 100, 100, 1000, 1000)
GrabURL('http://www.bing.com', 'file2.jpg', 100, 100, 1000, 1000)

cwd = os.getcwd()

outlook = win32.Dispatch('Outlook.Application')
Mail_Item = outlook.CreateItem(0)
Mail_Item.Recipients.Add('12345@qq.com')

Mail_Item.Subject ='test mail with images'
Mail_Item.BodyFormat = 2
Mail_Item.Attachments.Add(cwd + r'\file1.jpg')
Mail_Item.Attachments.Add(cwd + r'\file2.jpg')
Mail_Item.HtmlBody = '''
<html><body>
<a href=http://www.baidu.com>Baidu</a><br>
<img src=file1.jpg /><br>
<a href=http://www.bing.com>Bing</a><br>
<img src=file2.jpg /><br>
</body></html>
'''
Mail_Item.Display()
#Mail_Item.Send()
