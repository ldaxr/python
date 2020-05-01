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
    
GrabURL('http://www.baidu.com', 'baidu.jpg', 100, 100, 1000, 1000)
GrabURL('http://www.bing.com', 'bing.jpg', 100, 100, 1000, 1000)



outlook = win32.Dispatch('Outlook.Application')
Mail_Item = outlook.CreateItem(0)
Mail_Item.Recipients.Add('12345@qq.com')

Mail_Item.Subject ='test mail with images'
Mail_Item.BodyFormat = 2
Mail_Item.Attachments.Add(r'C:\Users\1325\Documents\baidu.jpg')
Mail_Item.Attachments.Add(r'C:\Users\1325\Documents\bing.jpg')
Mail_Item.Attachments.Add(r'C:\Users\1325\Documents\test.py')
Mail_Item.HtmlBody = '''
<html><body>
<a href=http://www.baidu.com>Baidu</a><br>
<img src=baidu.jpg /><br>
<a href=http://www.bing.com>Bing</a><br>
<img src=bing.jpg /><br>
</body></html>
'''
Mail_Item.Display()
Mail_Item.Send()
