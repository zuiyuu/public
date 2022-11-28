import os
import time
import yagmail
from imbox import Imbox
from PIL import ImageGrab

class mailTools:
    username = ''
    password = ''
    receiver = ''
    imapAdd = 'imap.163.com'
    smtpAdd = 'smtp.163.com'
    
    def __init__(self,un,pw) -> None:
        if (un == '' or pw == ''): 
            username = 'test@163.com'
            password = '你的授权码'
            receiver = 'test@qq.com'
        else:
            username = un
            password = pw
            receiver = un
        # yagmail.register(username, password)

    def send_mail(self, sender, to, contents,subjectText):
        if (sender == ''):
            sender = self.username
        if (to == '') :
            to = self.receiver
        # # smtp = yagmail.SMTP(user=sender, host='smtp.163.com')
        # smtp = yagmail.SMTP(user=sender, host=self.smtpAdd)
        # # smtp.send(to, subject='Remote Control', contents=contents)
        # smtp.send(to, subject=subjectText, contents=contents)
        
    def read_mail(self, username, password):
        # with Imbox('imap.163.com', username, password, ssl=True) as box:
        with Imbox(self.imapAdd, username, password, ssl=True) as box:
            all_msg = box.messages(unread=True)
            for uid, message in all_msg:
                # 如果是手机端发来的远程控制邮件
                if message.subject == 'Remote Control':
                    # 标记为已读
                    box.mark_seen(uid)
                    return message.body['plain'][0]
                
    def shutdown():
        os.system('shutdown -s -t 0')
