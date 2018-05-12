#-------------------------------------------------
#
# Project created by QtCreator 2018-05-04T15:29:19
#
#-------------------------------------------------

QT       += core gui axcontainer
CONFIG += qaxcontainer

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = excel_operation
TEMPLATE = app


SOURCES += main.cpp\
        mainwindow.cpp

HEADERS  += mainwindow.h

FORMS    += mainwindow.ui

CONFIG(debug, debug|release): DESTDIR +=  $$PWD/../bin/debug/win32
CONFIG(release, debug|release): DESTDIR +=  $$PWD/../bin/release/win32

        CONFIG(release, debug|release): LIBS +=  -lPsapi
        CONFIG(debug, debug|release): LIBS +=  -lPsapi
