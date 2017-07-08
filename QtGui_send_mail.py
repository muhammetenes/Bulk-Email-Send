import os
import time
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

from PyQt5 import QtWidgets, QtCore
import smtplib


class Window(QtWidgets.QWidget):
    def __init__(self):
        super(Window, self).__init__()

        self.text_label = QtWidgets.QLabel(self)
        self.text2_label = QtWidgets.QLabel(self)
        self.text = QtWidgets.QTextEdit(self)
        self.text2 = QtWidgets.QTextEdit(self)

        self.sender_mail_label = QtWidgets.QLabel(self)
        self.sender_mail = QtWidgets.QLineEdit(self)
        self.sender_password_label = QtWidgets.QLabel(self)
        self.sender_password = QtWidgets.QLineEdit(self)

        self.content_mail = QtWidgets.QTextEdit(self)
        self.subject_mail = QtWidgets.QLineEdit(self)

        self.content_label = QtWidgets.QLabel(self)
        self.subject_label = QtWidgets.QLabel(self)

        self.sender_server_label = QtWidgets.QLabel(self)
        self.server_port_label = QtWidgets.QLabel(self)

        self.sender_server = QtWidgets.QLineEdit(self)
        self.server_port = QtWidgets.QLineEdit(self)

        self.clr_btn = QtWidgets.QPushButton('Clear')
        self.save_btn = QtWidgets.QPushButton('Send')
        self.open_btn = QtWidgets.QPushButton('Open')

        self.pbar_label = QtWidgets.QLabel(self)
        self.pbar = QtWidgets.QProgressBar(self)
        self.pbar.setGeometry(30, 40, 200, 25)

        self.init_ui()

    def init_ui(self):
        v_layout = QtWidgets.QVBoxLayout()
        h_layout = QtWidgets.QHBoxLayout()
        h_layout_label = QtWidgets.QHBoxLayout()
        h_layout_edit = QtWidgets.QHBoxLayout()
        h_layout_server = QtWidgets.QHBoxLayout()
        h_layout_server_label = QtWidgets.QHBoxLayout()
        h_layout_sender = QtWidgets.QHBoxLayout()
        h_layout_sender_label = QtWidgets.QHBoxLayout()
        h_layout_pbar_label = QtWidgets.QHBoxLayout()
        h_layout_pbar = QtWidgets.QHBoxLayout()

        h_layout_pbar_label.addWidget(self.pbar_label)
        h_layout_pbar.addWidget(self.pbar)

        h_layout_sender.addWidget(self.sender_mail)
        h_layout_sender.addWidget(self.sender_password)

        self.sender_mail_label.setText('Sender Mail Address')
        self.sender_password_label.setText('Sender Mail Password')
        h_layout_sender_label.addWidget(self.sender_mail_label)
        h_layout_sender_label.addWidget(self.sender_password_label)

        h_layout.addWidget(self.open_btn)
        h_layout.addWidget(self.save_btn)
        h_layout.addWidget(self.clr_btn)

        self.server_port_label.setText('Server Port')
        self.server_port.setText('587')
        self.sender_server_label.setText('Mail Server')
        h_layout_server_label.addWidget(self.sender_server_label)
        h_layout_server_label.addWidget(self.server_port_label)

        h_layout_server.addWidget(self.sender_server)
        h_layout_server.addWidget(self.server_port)

        self.text_label.setText('Send To Mail')
        self.text2_label.setText('{{var1}}')

        h_layout_label.addWidget(self.text_label)
        h_layout_label.addWidget(self.text2_label)
        h_layout_edit.addWidget(self.text)
        h_layout_edit.addWidget(self.text2)

        v_layout.addLayout(h_layout_pbar_label)
        v_layout.addLayout(h_layout_pbar)

        v_layout.addLayout(h_layout_sender_label)
        v_layout.addLayout(h_layout_sender)

        v_layout.addLayout(h_layout_server_label)
        v_layout.addLayout(h_layout_server)

        self.content_label.setText('Content')
        v_layout.addWidget(self.content_label)
        v_layout.addWidget(self.content_mail)

        self.subject_label.setText('Subject')
        v_layout.addWidget(self.subject_label)
        v_layout.addWidget(self.subject_mail)

        v_layout.addLayout(h_layout_label)
        v_layout.addLayout(h_layout_edit)
        v_layout.addLayout(h_layout)

        self.open_btn.clicked.connect(self.open_text)
        self.save_btn.clicked.connect(self.send_text)
        self.clr_btn.clicked.connect(self.clr_text)

        self.setLayout(v_layout)
        self.setWindowTitle('Send Bulk Mail')

        self.show()

    def open_text(self):
        self.file_name = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', os.getenv('HOME'))
        if self.file_name.index(''):
            with open(self.file_name[0], 'r') as f:
                lines = f.readlines()
                count = 1

                for line in lines:
                    email = line.split(";")
                    email_1 = '{}-{}'.format(count, email[0].strip())
                    email_2 = '{}-{}'.format(count, email[1].strip())
                    self.text.append(email_1)
                    self.text2.append(email_2)
                    count = count + 1

                f.close()

        return self.file_name

    def send_text(self):
        try:
            mail = smtplib.SMTP(self.sender_server.text(), self.server_port.text())
            mail.ehlo_or_helo_if_needed()
            mail.starttls()
            mail.login(self.sender_mail.text(), self.sender_password.text().strip())
        except Exception as e:
            choice = QtWidgets.QMessageBox.question(self, 'Error', 'Mailing information is incorrect!',
                                                    QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok)

        if self.file_name.index(''):
            with open(self.file_name[0], 'r') as f:
                lines = f.readlines()
                lines_count = lines.__len__()
                percent = lines_count / 100
                percent_pbar = 0

                self.text.setDisabled(True)
                self.text2.setDisabled(True)
                self.content_mail.setDisabled(True)
                self.sender_mail.setDisabled(True)
                self.subject_mail.setDisabled(True)
                self.sender_password.setDisabled(True)
                self.sender_server.setDisabled(True)
                self.server_port.setDisabled(True)

                for_count = 1

                for line in lines:
                    email = line.split(";")
                    try:
                        content = self.content_mail.toPlainText().replace('{{var1}}', email[1])
                    except:
                        content = self.content_mail.toPlainText()

                    msg = MIMEMultipart()
                    msg['From'] = self.sender_mail.text()
                    msg['To'] = email[0]
                    msg['Date'] = formatdate(localtime=True)
                    msg['Subject'] = self.subject_mail.text()
                    text = MIMEText(content, 'plain')
                    msg.attach(text)

                    mail.sendmail(self.sender_mail.text(), email[0], msg.as_string())

                    for_count += 1

                    if round(for_count % percent, 1) == 0:
                        percent_pbar += 1
                        self.pbar.setValue(percent_pbar)
                    if for_count == lines_count:
                        percent_pbar = 100
                        self.pbar.setValue(percent_pbar)
                        self.text.setDisabled(False)
                        self.text2.setDisabled(False)
                        self.content_mail.setDisabled(False)
                        self.sender_mail.setDisabled(False)
                        self.subject_mail.setDisabled(False)
                        self.sender_password.setDisabled(False)
                        self.sender_server.setDisabled(False)
                        self.server_port.setDisabled(False)

                f.close()

        mail.close()

    def clr_text(self):
        self.text.clear()
        self.text2.clear()


app = QtWidgets.QApplication(sys.argv)
writer = Window()
sys.exit(app.exec_())
