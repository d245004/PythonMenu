# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'Menu.ui'
##
## Created by: Qt User Interface Compiler version 5.15.2
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide2.QtCore import *
from PySide2.QtGui import *
from PySide2.QtWidgets import *


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        if not Dialog.objectName():
            Dialog.setObjectName(u"Dialog")
        Dialog.resize(635, 391)
        font = QFont()
        font.setFamily(u"Agency FB")
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        Dialog.setFont(font)
        Dialog.setLayoutDirection(Qt.LeftToRight)
        self.layoutWidget = QWidget(Dialog)
        self.layoutWidget.setObjectName(u"layoutWidget")
        self.layoutWidget.setGeometry(QRect(20, 10, 321, 242))
        self.verticalLayout = QVBoxLayout(self.layoutWidget)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.pushButton = QPushButton(self.layoutWidget)
        self.pushButton.setObjectName(u"pushButton")

        self.verticalLayout.addWidget(self.pushButton)

        self.pushButton_2 = QPushButton(self.layoutWidget)
        self.pushButton_2.setObjectName(u"pushButton_2")

        self.verticalLayout.addWidget(self.pushButton_2)

        self.pushButton_3 = QPushButton(self.layoutWidget)
        self.pushButton_3.setObjectName(u"pushButton_3")

        self.verticalLayout.addWidget(self.pushButton_3)

        self.pushButton_4 = QPushButton(self.layoutWidget)
        self.pushButton_4.setObjectName(u"pushButton_4")

        self.verticalLayout.addWidget(self.pushButton_4)

        self.pushButton_5 = QPushButton(self.layoutWidget)
        self.pushButton_5.setObjectName(u"pushButton_5")

        self.verticalLayout.addWidget(self.pushButton_5)

        self.label = QLabel(self.layoutWidget)
        self.label.setObjectName(u"label")
        self.label.setAlignment(Qt.AlignCenter)

        self.verticalLayout.addWidget(self.label)

        self.lineEdit = QLineEdit(self.layoutWidget)
        self.lineEdit.setObjectName(u"lineEdit")
        self.lineEdit.setMaxLength(8)
        self.lineEdit.setAlignment(Qt.AlignCenter)

        self.verticalLayout.addWidget(self.lineEdit)

        self.pushButton_6 = QPushButton(self.layoutWidget)
        self.pushButton_6.setObjectName(u"pushButton_6")
        self.pushButton_6.setEnabled(True)

        self.verticalLayout.addWidget(self.pushButton_6)

        self.calendarWidget = QCalendarWidget(Dialog)
        self.calendarWidget.setObjectName(u"calendarWidget")
        self.calendarWidget.setGeometry(QRect(350, 10, 272, 241))
        font1 = QFont()
        font1.setFamily(u"\ub9d1\uc740 \uace0\ub515")
        font1.setPointSize(10)
        font1.setBold(True)
        font1.setWeight(75)
        self.calendarWidget.setFont(font1)
        self.calendarWidget.setGridVisible(True)
        self.calendarWidget.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.calendarWidget.setNavigationBarVisible(True)
        self.calendarWidget.setDateEditEnabled(True)
        self.pushButton_7 = QPushButton(Dialog)
        self.pushButton_7.setObjectName(u"pushButton_7")
        self.pushButton_7.setGeometry(QRect(20, 270, 321, 23))
        self.pushButton_8 = QPushButton(Dialog)
        self.pushButton_8.setObjectName(u"pushButton_8")
        self.pushButton_8.setGeometry(QRect(20, 300, 321, 23))
        self.pushButton_9 = QPushButton(Dialog)
        self.pushButton_9.setObjectName(u"pushButton_9")
        self.pushButton_9.setGeometry(QRect(20, 330, 321, 23))
        self.pushButton_10 = QPushButton(Dialog)
        self.pushButton_10.setObjectName(u"pushButton_10")
        self.pushButton_10.setGeometry(QRect(20, 360, 321, 23))

        self.retranslateUi(Dialog)

        QMetaObject.connectSlotsByName(Dialog)
    # setupUi

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(QCoreApplication.translate("Dialog", u"Python Work List", None))
#if QT_CONFIG(tooltip)
        self.pushButton.setToolTip(QCoreApplication.translate("Dialog", u"<html><head/><body><p>\ub9e4\ucd9c \ub2e4\uc6b4\ub85c\ub4dc \ud6c4 \uc791\uc5c5 \ud560 \uac83.</p></body></html>", None))
#endif // QT_CONFIG(tooltip)
        self.pushButton.setText(QCoreApplication.translate("Dialog", u"\ucee4\uba38\uc2a4 \ub9e4\ucd9c \uc785\ub825", None))
        self.pushButton_2.setText(QCoreApplication.translate("Dialog", u"\ucee4\uba38\uc2a4 \ub9e4\ucd9c \ube44\uad50", None))
        self.pushButton_3.setText(QCoreApplication.translate("Dialog", u"\ub85c\ub610 \ubc88\ud638 \ub9cc\ub4e4\uae30", None))
        self.pushButton_4.setText(QCoreApplication.translate("Dialog", u"\uc6d4\uc9d1\uacc4\ud45c \uc791\uc131", None))
#if QT_CONFIG(tooltip)
        self.pushButton_5.setToolTip(QCoreApplication.translate("Dialog", u"<html><head/><body><p>Group A, Group B, Grpup C, Group D, Group E \ud30c\uc77c\uc744 \ubc14\ud0d5\ud654\uba74\uc5d0 \ub9cc\ub4e4\uc5b4 \ub193\uace0\uc11c \uc2dc\uc791 \ud560 \uac83</p><p>\uc791\uc5c5\uc774 \ub05d\ub098\uba74 \uc77c\ubc18\ub9e4\ucd9c \ud569\uacc4\ud45c.xlsx \uc774 \ub9cc\ub4e4\uc5b4 \uc9c4\ub2e4</p></body></html>", None))
#endif // QT_CONFIG(tooltip)
        self.pushButton_5.setText(QCoreApplication.translate("Dialog", u"\uc77c\ubc18\ub9e4\ucd9c \ud569\uacc4 \uc870\ud68c", None))
        self.label.setText(QCoreApplication.translate("Dialog", u"\uacc4\uc0b0\uc11c \ubc1c\ud589 \uc77c\uc790\ub97c \uc785\ub825\ud558\uc138\uc694 ex: 20190101", None))
#if QT_CONFIG(tooltip)
        self.pushButton_6.setToolTip(QCoreApplication.translate("Dialog", u"<html><head/><body><p>\uc9c0\uae08\uc740 \ubc1c\ud589 \uc77c\uc790\ub97c \uc218\uc791\uc5c5\uc73c\ub85c \uc785\ub825\ud574\uc57c \ud55c\ub2e4</p><p>input \ud568\uc218\uac00 error\ub97c \uc720\ubc1c\ud55c\ub2e4</p></body></html>", None))
#endif // QT_CONFIG(tooltip)
        self.pushButton_6.setText(QCoreApplication.translate("Dialog", u"\uc6d4 \uc138\uae08\uacc4\uc0b0\uc11c \ubc1c\ud589", None))
        self.pushButton_7.setText(QCoreApplication.translate("Dialog", u"\uacac\uc801\uc11c \ubcc0\ud658", None))
        self.pushButton_8.setText(QCoreApplication.translate("Dialog", u"\uc804\ubb38\uc810 \uc870\uce58\uac74 \ud604\ub300,\uae30\uc544 \ud569\uce58\uae30", None))
        self.pushButton_9.setText(QCoreApplication.translate("Dialog", u"MOBI 6AM \uc870\ud68c", None))
        self.pushButton_10.setText(QCoreApplication.translate("Dialog", u"ABC delete \ucd94\ucd9c", None))
    # retranslateUi

