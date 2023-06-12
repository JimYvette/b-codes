QT       += core gui
QT       += core gui sql
QT       += charts

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++11
INCLUDEPATH +=C:\eigen
CONFIG += warn_off
QMAKE_CXXFLAGS += -Wall
QMAKE_CXXFLAGS += -Wno-commen
# The following define makes your compiler emit warnings if you use
# any Qt feature that has been marked deprecated (the exact warnings
# depend on your compiler). Please consult the documentation of the
# deprecated API in order to know how to port your code away from it.
DEFINES += QT_DEPRECATED_WARNINGS

# You can also make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
# You can also select to disable deprecated APIs only up to a certain version of Qt.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

# 版本信息
VERSION = 4.0.2.777

# 图标
RC_ICONS = pic1.ico

# 产品名称
QMAKE_TARGET_PRODUCT = "Qt Creator"

# 文件说明
QMAKE_TARGET_DESCRIPTION = "Qt Creator based on Qt 5.14.2 (MinGW, 64 bit)"

# 版权信息
QMAKE_TARGET_COPYRIGHT = "Copyright 2008-2022 The Qt Company Ltd. All rights reserved."

# 中文（简体）
RC_LANG = 0x0004


SOURCES += \
    form.cpp \
    main.cpp

HEADERS += \
    form.h

FORMS += \
    form.ui

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target

RESOURCES += \
    images.qrc
QT += axcontainer
