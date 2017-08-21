#-------------------------------------------------
#
# Project created by QtCreator 2017-08-13T08:33:50
#
#-------------------------------------------------

QT       += core gui
QT += network axcontainer

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = attendanceget
TEMPLATE = app
DESTDIR = ./attendanceget_bin


SOURCES += main.cpp\
        mainwindow.cpp

HEADERS  += mainwindow.h \
    common.h

FORMS    += mainwindow.ui
