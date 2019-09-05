"""Client FOR WORKING WITH MAIL SERVERS POP3 AND SMTP PROTOCOLS
Using API Kivi"""

import sys
import smtplib
import os
import poplib
import email
from email.header import decode_header
from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.config import Config
from kivy.adapters.listadapter import ListAdapter
from kivy.uix.listview import ListItemButton, ListView
from kivy.uix.listview import ListItemButton
from functools import partial
from os.path import basename
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate
from kivy.uix.popup import Popup
from  kivy.uix.filechooser import FileChooserListView

Config.set('graphics', 'resizable', 0)
Config.set('graphics', 'width', 800)
Config.set('graphics', 'height', 640)

class EmailApp(App):

    def clear_all(self):
        child = self.root.children[:]
        for i in child :self.root.remove_widget(i)

    def send_message(self,instance):

        def send_email_with_attachment(host, port, subject, to_addr, from_addr, body_text, login, password,
                                       file_to_attach):
            header = 'Content-Disposition', 'attachment; filename="%s"' % basename(file_to_attach)

            msg = MIMEMultipart()
            msg["From"] = from_addr
            msg["Subject"] = subject
            msg["Date"] = formatdate(localtime=True)

            if body_text:
                msg.attach(MIMEText(body_text))

            msg["To"] = to_addr
            attachment = MIMEBase('application', "octet-stream")

            if file_to_attach != '':
                try:
                    with open(file_to_attach, "rb") as fh:
                        data = fh.read()

                    attachment.set_payload(data)
                    encoders.encode_base64(attachment)
                    attachment.add_header(*header)
                    msg.attach(attachment)

                except IOError:
                    msg = "Error opening attachment file %s" % file_to_attach
                    print(msg)
                    sys.exit(1)

            server = smtplib.SMTP_SSL(host, )
            server.login(login, password)
            server.sendmail(from_addr, [to_addr], msg.as_string())
            server.quit()

        global login, password, filep, server
        self.server_input_send.focus = False
        self.subject_input.focus = False
        self.adress_input.focus = False
        self.body_input.focus = False
        host = self.server_input_send.text[7:]
        subject = self.subject_input.text[8:]
        to_addr = self.adress_input.text[3:]
        from_addr = login + '@' +server[server.find('.')+1:]
        body_text = self.body_input.text[5:]
        print(host, subject, to_addr, from_addr, body_text, login, password, filep)

        try:
            send_email_with_attachment(host, port, subject, to_addr, from_addr, body_text, login, password, filep)
        except:
            self.server_input_send.focus = True
            self.server_input_send.text = 'INVALID PARAMETERS'
        else:
            self.rebuild(instance)

    def connection(self):
        def do_decode_header(header):

            header_parts = decode_header(header)

            res = []
            for decoded_string, encoding in header_parts:
                if encoding:
                    decoded_string = decoded_string.decode(encoding)
                elif isinstance(decoded_string, bytes):
                    decoded_string = decoded_string.decode("ascii")
                res.append(decoded_string)

            return "".join(res)

        def get_part_info(part):
            encoding = part.get_content_charset()

            if not encoding:
                encoding = sys.stdout.encoding
            mime = part.get_content_type()
            message = part.get_payload(decode=True).decode(encoding, errors="ignore").strip()

            return message, encoding, mime

        def get_message_info(message):

            message_text, encoding, mime = "Нет тела сообщения", "-", "-"

            if message.is_multipart():
                for part in message.walk():
                    if part.get_content_type() in ("text/html", "text/plain"):
                        message_text, encoding, mime = get_part_info(part)
                        break
            else:
                message_text, encoding, mime = get_part_info(message)

            return message_text, encoding, mime

        if __name__ == "__main__":

            global server, port, login, password, filepath

            with open(__file__ + "_data.txt", "w", encoding="utf-8") as fh:
                try:
                    mailserver = poplib.POP3_SSL(server, port)
                    mailserver.user(login)
                    mailserver.pass_(password)

                    stat = mailserver.stat()
                    print("\nВсего писем: {}, объем: {:.2f} Мб.".
                          format(stat[0], stat[1] / 2 ** 20), file=fh)
                    self.lb1.text=("\nВсего писем: {}, объем: {:.2f} Мб.".
                          format(stat[0], stat[1] / 2 ** 20))

                    response, messages, octets = mailserver.list()
                    print("\nСписок писем: ", response.decode(), file=fh)

                    global list_of_messages
                    list_item = []
                    list_of_messages = [''] * (len(messages)+1)
                    for message_num in range(len(messages)):
                        message_num = len(messages) - message_num

                        response, raw_message, octets = mailserver.retr(message_num)
                        if response.decode()[0:3] != "+OK":
                            print("Не удалось получить письмо №", message_num, file=fh)
                            continue

                        raw_message = b"\n".join(raw_message)
                        message = email.message_from_bytes(raw_message)

                        text, encoding, mime = get_message_info(message)

                        list_of_messages[message_num] = ("\n  - от: '{}'"
                              "\n  - тема: '{}'"
                              "\n  - идентификатор сообщения: "
                              "\n   '{}'"
                              "\n  - дата: '{}'"
                              "\n  - объем: {:.2f} Кб."
                              "\n  - содержимое:\n{} ".
                              format(do_decode_header(message["From"]),
                                     do_decode_header(message["Subject"]),
                                     do_decode_header(message["Message-Id"]),
                                     message["Date"],
                                     octets / 2 ** 10,
                                     text))
                        list_item.append(str(message_num) + '. ' + do_decode_header(message["From"])+ do_decode_header(message["Subject"]))
                    print("Информация о почтовом ящике и письма прочитаны и сохранены.")
                except poplib.error_proto as err:
                    print("Возникла следующая ошибка:", err)
                finally:
                    pass
                    #mailserver.quit()
                self.list_adapter.data = list_item

    def attachment(self):
        global server, port, login, password, filepath
        def get_attachment(message, filepath):
            if message.is_multipart():
                for part in message.walk():

                    if part.get_content_maintype() == 'multipart':
                        continue

                    if part.get('Content-Disposition') is None:
                        continue

                    filename = part.get_filename()
                    if not (filename): filename = "test.txt"
                    filename = decode_header(filename)[0][0]
                    fp = open(os.path.join(filepath) + '/' + filename, 'wb')
                    fp.write(part.get_payload(decode=1))
                    fp.close
        mailserver = poplib.POP3_SSL(server, port)
        mailserver.user(login)
        mailserver.pass_(password)
        response, raw_message, octets = mailserver.retr(int(text_mes[0]))
        raw_message = b"\n".join(raw_message)
        message = email.message_from_bytes(raw_message)

        get_attachment(message, filepath)
        pass

    def del_mes(self,instance):
        global server, port, login, password, filepath
        mailserver = poplib.POP3_SSL(server, port)
        mailserver.user(login)
        mailserver.pass_(password)
        mailserver.dele(int(text_mes[0]))
        mailserver.quit()
        self.rebuild(instance)

    def show_mes(self):
        global text_mes
        global list_of_messages
        mes = list_of_messages[int(text_mes[0])]
        mes = mes.split('\n')
        listitem1 = ListItemButton
        self.list_adapter = ListAdapter(data=mes, cls=listitem1, allow_empty_selection=False)
        self.list_view = ListView(adapter=self.list_adapter)
        listitem1.background_color=[0,0,0,1]
        listitem1.font_size=17
        listitem1.halign = 'left'
        listitem1.text_size = (700,0)
        self.clear_all()
        b2 = BoxLayout(orientation='horizontal', padding=25, size_hint=(1, .2))
        b2.add_widget(Button(text='Back', on_release=self.rebuild, on_press=self.update))
        b2.add_widget(Button(text='Delete', on_release=self.del_mes))
        self.root.add_widget(b2)

        self.root.add_widget(self.list_view)
        self.attachment()
        pass

    class list_but(ListItemButton):
        def show_mes(self):
            global text_mes
            text_mes = (self.text)
            print(self.text)

        on_press = show_mes
        pass

    def on_focus(self, instance, value):
        textr = instance.id
        if value:
            instance.text = instance.text[len(textr):]
        else:
            instance.text = textr + instance.text

    def dismiss_popup(self,instance):
        self._popup.dismiss()

    def save(self, path, selection, instance):
        global filep
        filep = self.filechooser.selection and self.filechooser.selection[0] or '1'
        print(filep)
        self.dismiss_popup(instance)

    def show_popup(self,instance):
        b2 = BoxLayout(orientation='vertical', padding=25)
        self.filechooser = (FileChooserListView( path=r'C:\tmp'))
        b2.add_widget(self.filechooser)
        b3 = BoxLayout(orientation='horizontal', padding=25, size_hint=(1, .2))
        b3.add_widget(Button(text='Back', on_release=self.dismiss_popup))
        choose_but = Button(text='Choose')
        b3.add_widget(choose_but)
        choose_but.bind(on_release=partial(self.save, self.filechooser.path,self.filechooser.selection))
        b2.add_widget(b3)
        content = b2
        self._popup = Popup(title="Choose file", content=content, size_hint=(0.9, 0.9))
        self._popup.open()

    def rebuild_send(self, instance):
        self.clear_all()
        global filep
        filep = ''
        self.server_input_send = (TextInput(multiline=False, text = 'Server:', id = 'Server:', size_hint=(1, .1)))
        self.server_input_send.bind(focus=self.on_focus)
        self.port_input_send = (TextInput(multiline=False, text = 'Port:', id = 'Port:', size_hint=(1, .1)))
        self.port_input_send.bind(focus=self.on_focus)
        self.adress_input = (TextInput(multiline=False, text = 'To:',id = 'To:', size_hint=(1, .1)))
        self.adress_input.bind(focus=self.on_focus)
        self.subject_input = (TextInput(multiline=False, text = 'Subject:', id = 'Subject:', size_hint=(1, .1)))
        self.subject_input.bind(focus=self.on_focus)
        self.body_input = (TextInput(text = 'Body:',id = 'Body:', size_hint=(1, .5)))
        self.body_input.bind(focus=self.on_focus)

        b2 = BoxLayout(orientation='horizontal', padding=25, size_hint=(1, .2))
        b2.add_widget(Button(text='Back', on_release=self.rebuild, on_press=self.update))
        b2.add_widget(Button(text='Attachment', on_release=self.show_popup))
        self.root.add_widget(b2)
        self.root.add_widget(self.server_input_send)
        self.root.add_widget(self.port_input_send)
        self.root.add_widget(self.adress_input)
        self.root.add_widget(self.subject_input)
        self.root.add_widget(self.body_input)
        self.root.add_widget(Button(text='Send', on_press=self.send_message, size_hint=(1, .1)))

    def update(self,instance):
        try:
            self.clear_all()
            self.lb1 = (Label(text='Connecting...', font_size=50, halign='left', size_hint=(1, .1)))
            self.root.add_widget(self.lb1)
        except ValueError:
            self.port_input.text = 'MUST BE INTEGER'
        except:
            self.server_input.text = 'INVALID PARAMETERS'
        finally:
            self.rebuild(instance)

    def rebuild(self, instance):
        try:
            self.clear_all()
            self.lb1 = (Label(text='Information', font_size=20, halign='left', size_hint=(1, .1)))
            self.root.add_widget(self.lb1)

            b2 = BoxLayout(orientation='horizontal', padding=25, size_hint=(1, .2))
            b2.add_widget(Button(text='Write message', on_press=self.rebuild_send))
            b2.add_widget(Button(text='Update', on_release=self.rebuild, on_press = self.update))
            self.root.add_widget(b2)

            self.list_button = self.list_but#ListItemButton
            self.list_adapter = ListAdapter(data = [], cls=self.list_button,allow_empty_selection=False)

            self.list_view = ListView(adapter=self.list_adapter)
            self.list_button.background_color = [1, 1, .9, 1]
            self.list_button.font_size = 14
            self.list_button.halign = 'left'
            self.list_button.text_size = (700, 0)

            self.root.add_widget(self.list_view)
            self.list_button.on_release = self.show_mes
            self.connection()
        except ValueError:
            self.port_input.text = 'MUST BE INTEGER'
        except:
            self.server_input.text = 'INVALID PARAMETERS'

    def parametrs(self, instance):
        try:
            global server, port, login, password, filepath
            server = self.server_input.text
            port = int(self.port_input.text)
            login = self.login_input.text
            password = self.pass_input.text
            filepath = self.location_input.text
            mailserver = poplib.POP3_SSL(server, port)
            mailserver.user(login)
            mailserver.pass_(password)
            self.update(instance)
        except ValueError:
            self.port_input.text = 'MUST BE INTEGER'
        except:
            self.server_input.text = 'INVALID PARAMETERS'

    def build(self):
        b1 = BoxLayout(orientation='vertical', padding=25)
        g1 = GridLayout(cols=2, spacing=60, padding=25, size_hint=(1, .8))

        self.lb1 = (Label(text='Login', font_size=40, halign='left', size_hint=(.2, .1)))
        self.login_input = (TextInput(multiline=False))

        self.lb2 = (Label(text='Password', font_size=40, halign='right', size_hint=(.2, .1)))
        self.pass_input = (TextInput(multiline=False))

        self.lb3 = (Label(text='Server', font_size=40, halign='right', size_hint=(.2, .1)))
        self.server_input = (TextInput(multiline=False))

        self.lb4 = (Label(text='Port', font_size=40, halign='right', size_hint=(.2, .1)))
        self.port_input = (TextInput(multiline=False))

        self.lb5 = (Label(text='Location', font_size=40, halign='right', size_hint=(.2, .1)))
        self.location_input = (TextInput(multiline=False))

        g1.add_widget(self.lb1)
        g1.add_widget(self.login_input)
        g1.add_widget(self.lb2)
        g1.add_widget(self.pass_input)
        g1.add_widget(self.lb3)
        g1.add_widget(self.server_input)
        g1.add_widget(self.lb4)
        g1.add_widget(self.port_input)
        g1.add_widget(self.lb5)
        g1.add_widget(self.location_input)
        b1.add_widget(g1)
        b1.add_widget(Button(text='Connect', on_press = self.parametrs, size_hint=(1, .1)))
        return b1

if __name__ == "__main__":
    EmailApp().run()
