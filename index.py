import datetime

from PyQt5 import QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
import mysql.connector
from PyQt5.uic import loadUiType
from xlrd import *
from xlsxwriter import *

# from PyQt5.QtWidgets import QDialog

ui, _ = loadUiType('library.ui')
login, _ = loadUiType('login.ui')


class Login(QWidget, login):
    def __init__(self):
        QWidget.__init__(self)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.handle_login)
        self.qdark_theme()

    def qdark_theme(self):
        style = open('themes/qdark.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def handle_login(self):

        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        username = self.lineEdit.text()
        password = self.lineEdit_2.text()

        sql = '''SELECT * FROM users'''
        self.cur.execute(sql)
        data = self.cur.fetchall()
        for info in data:
            if username == info[1] and password == info[3]:

                self.window2 = MainApp()
                self.close()
                self.window2.show()
            else:
                self.label.setText('Enter Correct User Name and Password')


class MainApp(QMainWindow, ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.handle_ui_changes()
        self.handle_buttons()
        self.qdark_theme()

        self.show_author()
        self.show_category()
        self.show_publisher()

        self.show_category_combobox()
        self.show_author_combobox()
        self.show_publisher_combobox()

        self.show_all_clients()
        self.show_all_books()

        self.show_all_tasks()

    def handle_ui_changes(self):
        self.hide_themes()
        self.tabWidget.tabBar().setVisible(False)

    def handle_buttons(self):
        self.pushButton_5.clicked.connect(self.show_themes)
        self.pushButton_26.clicked.connect(self.hide_themes)

        self.pushButton.clicked.connect(self.open_day_to_day_tab)
        self.pushButton_2.clicked.connect(self.open_books_tab)
        self.pushButton_3.clicked.connect(self.open_users_tab)
        self.pushButton_4.clicked.connect(self.open_settings_tab)
        self.pushButton_31.clicked.connect(self.open_clients_tab)

        self.pushButton_7.clicked.connect(self.add_new_book)
        self.pushButton_10.clicked.connect(self.search_books)
        self.pushButton_9.clicked.connect(self.edit_books)
        self.pushButton_14.clicked.connect(self.delete_books)

        self.pushButton_19.clicked.connect(self.add_category)
        self.pushButton_20.clicked.connect(self.add_author)
        self.pushButton_21.clicked.connect(self.add_publisher)

        self.pushButton_15.clicked.connect(self.add_new_user)
        self.pushButton_16.clicked.connect(self.login)
        self.pushButton_18.clicked.connect(self.edit_user)

        self.pushButton_22.clicked.connect(self.dark_orange_theme)
        self.pushButton_23.clicked.connect(self.dark_blue_theme)
        self.pushButton_24.clicked.connect(self.dark_gray_theme)
        self.pushButton_25.clicked.connect(self.qdark_theme)

        self.pushButton_8.clicked.connect(self.add_new_client)
        self.pushButton_12.clicked.connect(self.search_client)
        self.pushButton_11.clicked.connect(self.edit_client)
        self.pushButton_27.clicked.connect(self.delete_client)

        self.pushButton_6.clicked.connect(self.handle_day_operations)

        self.pushButton_29.clicked.connect(self.export_day_opps)
        self.pushButton_13.clicked.connect(self.export_books)
        self.pushButton_28.clicked.connect(self.export_clients)

    def show_themes(self):
        self.groupBox_3.show()

    def hide_themes(self):
        self.groupBox_3.hide()

    #############################
    ####### opening tabs ########

    def open_day_to_day_tab(self):
        self.tabWidget.setCurrentIndex(0)

    def open_books_tab(self):
        self.tabWidget.setCurrentIndex(1)

    def open_clients_tab(self):
        self.tabWidget.setCurrentIndex(2)

    def open_users_tab(self):
        self.tabWidget.setCurrentIndex(3)

    def open_settings_tab(self):
        self.tabWidget.setCurrentIndex(4)

    ############################################
    #######     day-to-day operation    ########
    def handle_day_operations(self):
        client_name = self.lineEdit_36.text()
        book_title = self.lineEdit.text()
        type = self.comboBox.currentText()
        days_number = self.comboBox_2.currentText()
        today_date = datetime.date.today()
        to_date = today_date + datetime.timedelta(days=int(days_number))

        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''
            INSERT INTO dayoperations(book_name, client, type, days, date, to_date)
            VALUES (%s, %s, %s, %s, %s, %s)
        ''', (book_title, client_name, type, days_number, today_date, to_date))

        self.db.commit()
        self.statusBar().showMessage('Task Added')
        self.show_all_tasks()

    def show_all_tasks(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''
            SELECT book_name, client, type, date, to_date FROM dayoperations
        ''')

        data = self.cur.fetchall()

        self.tableWidget.setRowCount(100)
        tablerow = 0
        for row in data:
            self.tableWidget.setItem(tablerow, 0, QtWidgets.QTableWidgetItem(row[0]))
            self.tableWidget.setItem(tablerow, 1, QtWidgets.QTableWidgetItem(row[1]))
            self.tableWidget.setItem(tablerow, 2, QtWidgets.QTableWidgetItem(row[2]))
            self.tableWidget.setItem(tablerow, 3, QtWidgets.QTableWidgetItem(str(row[3])))
            self.tableWidget.setItem(tablerow, 4, QtWidgets.QTableWidgetItem(str(row[4])))

            tablerow += 1

    #############################
    #######     books    ########
    def show_all_books(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT book_title, book_description, book_code, book_category, book_author, 
        book_publisher, book_price FROM book''')
        data = self.cur.fetchall()

        self.tableWidget_5.setRowCount(10)
        tablerow = 0
        for row in data:
            self.tableWidget_5.setItem(tablerow, 0, QtWidgets.QTableWidgetItem(row[0]))
            self.tableWidget_5.setItem(tablerow, 1, QtWidgets.QTableWidgetItem(row[1]))
            self.tableWidget_5.setItem(tablerow, 2, QtWidgets.QTableWidgetItem(row[2]))
            self.tableWidget_5.setItem(tablerow, 3, QtWidgets.QTableWidgetItem(str(row[3])))
            self.tableWidget_5.setItem(tablerow, 4, QtWidgets.QTableWidgetItem(str(row[4])))
            self.tableWidget_5.setItem(tablerow, 5, QtWidgets.QTableWidgetItem(str(row[5])))
            self.tableWidget_5.setItem(tablerow, 6, QtWidgets.QTableWidgetItem(str(row[6])))

            tablerow += 1

    def add_new_book(self):

        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        book_title = self.lineEdit_2.text()
        book_description = self.textEdit.toPlainText()
        book_code = self.lineEdit_3.text()
        book_category = self.comboBox_3.currentText()
        book_author = self.comboBox_4.currentText()
        book_publisher = self.comboBox_5.currentText()
        book_price = self.lineEdit_4.text()

        self.cur.execute('''
            INSERT INTO book (book_title, book_description, book_code, book_category, book_author, book_publisher, book_price)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            ''', (book_title, book_description, book_code, book_category, book_author, book_publisher, book_price))

        self.db.commit()
        self.statusBar().showMessage('New Book Added')

        self.lineEdit_2.setText('')
        self.textEdit.setPlainText('')
        self.lineEdit_3.setText('')
        self.comboBox_3.setCurrentIndex(0)
        self.comboBox_4.setCurrentIndex(0)
        self.comboBox_5.setCurrentIndex(0)
        self.lineEdit_4.setText('')

        self.show_all_books()

    def search_books(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        book_title = self.lineEdit_12.text()

        sql = '''SELECT * FROM book WHERE book_title = %s'''
        self.cur.execute(sql, [(book_title)])

        data = self.cur.fetchone()

        self.lineEdit_20.setText(data[1])
        self.textEdit_2.setPlainText(data[2])
        self.lineEdit_13.setText(data[3])
        self.comboBox_13.setCurrentText(data[4])
        self.comboBox_14.setCurrentText(data[5])
        self.comboBox_12.setCurrentText(data[6])
        self.lineEdit_11.setText(str(data[7]))

    def edit_books(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        book_title = self.lineEdit_20.text()
        book_description = self.textEdit_2.toPlainText()
        book_code = self.lineEdit_13.text()
        book_category = self.comboBox_13.currentText()
        book_author = self.comboBox_14.currentText()
        book_publisher = self.comboBox_12.currentText()
        book_price = self.lineEdit_11.text()

        search_book_title = self.lineEdit_12.text()

        self.cur.execute('''
            UPDATE book SET book_title=%s ,book_description=%s ,book_code=%s ,book_category=%s ,book_author=%s ,book_publisher=%s ,book_price=%s WHERE book_title=%s
        ''', (book_title, book_description, book_code, book_category, book_author, book_publisher, book_price,
              search_book_title))

        self.db.commit()
        self.statusBar().showMessage('Book Updated')
        self.show_all_books()

    def delete_books(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        book_title = self.lineEdit_12.text()

        warning = QMessageBox.warning(self, 'Delete Book', 'Are you sure you want to delete this book',
                                      QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes:
            self.cur.execute('''DELETE FROM book WHERE book_title = %s''', [(book_title)])
            self.db.commit()
            self.statusBar().showMessage('Book Deleted')
            self.show_all_books()

    #############################
    #######    clients   ########
    def show_all_clients(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT client_name, client_email, client_nic FROM clients''')
        data = self.cur.fetchall()

        self.tableWidget_6.setRowCount(10)
        tablerow = 0
        for row in data:
            self.tableWidget_6.setItem(tablerow, 0, QtWidgets.QTableWidgetItem(row[0]))
            self.tableWidget_6.setItem(tablerow, 1, QtWidgets.QTableWidgetItem(row[1]))
            self.tableWidget_6.setItem(tablerow, 2, QtWidgets.QTableWidgetItem(row[2]))

            tablerow += 1

    def add_new_client(self):
        client_name = self.lineEdit_5.text()
        client_email = self.lineEdit_6.text()
        client_nic = self.lineEdit_7.text()

        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''
            INSERT INTO clients (client_name, client_email, client_nic)
            VALUES (%s, %s, %s)
        ''', (client_name, client_email, client_nic))

        self.db.commit()
        self.statusBar().showMessage('New Client Added')
        self.db.close()
        self.show_all_clients()

    def search_client(self):
        client_nic = self.lineEdit_15.text()
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT * FROM clients WHERE client_nic = %s ''', [(client_nic)])
        data = self.cur.fetchone()

        self.lineEdit_34.setText(data[1])
        self.lineEdit_16.setText(data[2])
        self.lineEdit_14.setText(data[3])

    def edit_client(self):
        client_prev_nic = self.lineEdit_15.text()

        client_name = self.lineEdit_34.text()
        client_email = self.lineEdit_16.text()
        client_nic = self.lineEdit_14.text()

        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''
            UPDATE clients SET client_name = %s, client_email = %s, client_nic = %s WHERE client_nic = %s
        ''', (client_name, client_email, client_nic, client_prev_nic))

        self.db.commit()
        self.statusBar().showMessage('Client Data Updated')
        self.show_all_clients()

    def delete_client(self):
        client_prev_nic = self.lineEdit_15.text()

        warning_message = QMessageBox.warning(self, "Delete Clients", "Are you sure you want to delete this client",
                                              QMessageBox.Yes | QMessageBox.No)

        if warning_message == QMessageBox.Yes:
            self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
            self.cur = self.db.cursor()

            sql = '''DELETE FROM clients WHERE client_nic = %s'''

            self.cur.execute(sql, [client_prev_nic])

            self.db.commit()
            self.db.close()
            self.statusBar().showMessage('Client Deleted')
            self.show_all_clients()

    #############################
    #######     users    ########

    def add_new_user(self):

        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        username = self.lineEdit_21.text()
        email = self.lineEdit_22.text()
        password = self.lineEdit_23.text()
        password2 = self.lineEdit_24.text()

        if password == password2:

            self.cur.execute('''
            INSERT INTO users(user_name , user_email , user_password)
            VALUES (%s , %s , %s)
            ''', (username, email, password))

            self.db.commit()
            self.statusBar().showMessage('New User Added')

        else:
            self.label_9.setText('Please type same password twice')

    def login(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        username = self.lineEdit_26.text()
        password = self.lineEdit_25.text()

        sql = '''SELECT * FROM users'''
        self.cur.execute(sql)
        data = self.cur.fetchall()
        for info in data:
            if username == info[1] and password == info[3]:
                self.statusBar().showMessage('Valid Username & Password')
                self.groupBox_4.setEnabled(True)

                self.lineEdit_28.setText(info[1])
                self.lineEdit_29.setText(info[2])
                self.lineEdit_27.setText(info[3])

    def edit_user(self):
        username = self.lineEdit_28.text()
        email = self.lineEdit_29.text()
        password = self.lineEdit_27.text()
        password2 = self.lineEdit_30.text()

        original_name = self.lineEdit_26.text()

        if password == password2:
            self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
            self.cur = self.db.cursor()

            self.cur.execute('''
                UPDATE users SET user_name = %s, user_email = %s, user_password = %s WHERE user_name = %s
            ''', (username, email, password, original_name))

            self.db.commit()
            self.statusBar().showMessage('User Data Updated Successfully')


    #############################
    #######   settings   ########

    def add_category(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        category_name = self.lineEdit_31.text()

        self.cur.execute('''
            INSERT INTO category (category_name) VALUES (%s)
        ''', (category_name,))

        self.db.commit()
        self.statusBar().showMessage('New Category Added')
        self.lineEdit_31.setText('')
        self.show_category()
        self.show_category_combobox()

    def show_category(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()
        self.cur.execute('''SELECT category_name FROM category''')
        data = self.cur.fetchall()

        self.tableWidget_2.setRowCount(10)
        tablerow = 0
        for row in data:
            self.tableWidget_2.setItem(tablerow, 0, QtWidgets.QTableWidgetItem(row[0]))
            tablerow += 1

    def add_author(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        author_name = self.lineEdit_32.text()

        self.cur.execute('''
            INSERT INTO authors (author_name) VALUES (%s)
                ''', (author_name,))

        self.db.commit()
        self.lineEdit_32.setText('')
        self.statusBar().showMessage('New Author Added')

        self.show_author()
        self.show_author_combobox()

    def show_author(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()
        self.cur.execute('''SELECT author_name FROM authors''')
        data = self.cur.fetchall()

        self.tableWidget_3.setRowCount(10)
        tablerow = 0
        for row in data:
            self.tableWidget_3.setItem(tablerow, 0, QtWidgets.QTableWidgetItem(row[0]))
            tablerow += 1

    def add_publisher(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        publisher_name = self.lineEdit_33.text()

        self.cur.execute('''
             INSERT INTO publisher (publisher_name) VALUES (%s)
       ''', (publisher_name,))

        self.db.commit()
        self.lineEdit_33.setText('')
        self.statusBar().showMessage('New Publisher Added')

        self.show_publisher()
        self.show_publisher_combobox()

    def show_publisher(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()
        self.cur.execute('''SELECT publisher_name FROM publisher''')
        data = self.cur.fetchall()

        self.tableWidget_4.setRowCount(10)
        tablerow = 0
        for row in data:
            self.tableWidget_4.setItem(tablerow, 0, QtWidgets.QTableWidgetItem(row[0]))
            tablerow += 1

    ############################################
    #######  show settings data in UI   ########
    def show_category_combobox(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()
        self.cur.execute('''SELECT category_name FROM category''')
        data = self.cur.fetchall()
        self.comboBox_3.clear()

        for category in data:
            self.comboBox_3.addItem(category[0])
            self.comboBox_13.addItem(category[0])

    def show_author_combobox(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()
        self.cur.execute('''SELECT author_name FROM authors''')
        data = self.cur.fetchall()
        self.comboBox_4.clear()

        for author in data:
            self.comboBox_4.addItem(author[0])
            self.comboBox_14.addItem(author[0])

    def show_publisher_combobox(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()
        self.cur.execute('''SELECT publisher_name FROM publisher''')
        data = self.cur.fetchall()
        self.comboBox_5.clear()

        for publisher in data:
            self.comboBox_5.addItem(publisher[0])
            self.comboBox_12.addItem(publisher[0])

    ############################
    #######  Export Data #######
    def export_day_opps(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''
            SELECT book_name, client, type, date, to_date FROM dayoperations
        ''')

        data = self.cur.fetchall()
        wb = Workbook('day_opps.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0, 0, 'Book Title')
        sheet1.write(0, 1, 'Client Name')
        sheet1.write(0, 2, 'Type')
        sheet1.write(0, 3, 'From - date')
        sheet1.write(0, 4, 'To - date')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        self.statusBar().showMessage('Report Created Successfully')

    def export_books(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT book_title, book_description, book_code, book_category, book_author, 
                book_publisher, book_price FROM book''')
        data = self.cur.fetchall()

        wb = Workbook('all_books.xlsx')
        sheet1 = wb.add_worksheet()
        sheet1.write(0, 0, 'Book Title')
        sheet1.write(0, 1, 'Description')
        sheet1.write(0, 2, 'Code')
        sheet1.write(0, 3, 'Category')
        sheet1.write(0, 4, 'Author')
        sheet1.write(0, 5, 'Publisher')
        sheet1.write(0, 6, 'Price')


        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        self.statusBar().showMessage('Report Created Successfully')

    def export_clients(self):
        self.db = mysql.connector.connect(host='localhost', user='root', password='@UsmanDB', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT client_name, client_email, client_nic FROM clients''')
        data = self.cur.fetchall()

        wb = Workbook('all_clients.xlsx')
        sheet1 = wb.add_worksheet()
        sheet1.write(0, 0, 'Client Name')
        sheet1.write(0, 1, 'Email')
        sheet1.write(0, 2, 'NIC')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        self.statusBar().showMessage('Report Created Successfully')

    ############################
    #######  UI Themes  ########

    def dark_blue_theme(self):
        style = open('themes/darkblue.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def dark_gray_theme(self):
        style = open('themes/darkgray.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def dark_orange_theme(self):
        style = open('themes/darkorange.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def qdark_theme(self):
        style = open('themes/qdark.css', 'r')
        style = style.read()
        self.setStyleSheet(style)


def main():
    app = QApplication(sys.argv)
    window = Login()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()
