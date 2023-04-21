/********************************************************************************
** Form generated from reading UI file 'excelOperationDemo.ui'
**
** Created by: Qt User Interface Compiler version 5.14.2
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_EXCELOPERATIONDEMO_H
#define UI_EXCELOPERATIONDEMO_H

#include <QtCore/QVariant>
#include <QtWidgets/QApplication>
#include <QtWidgets/QMainWindow>
#include <QtWidgets/QMenuBar>
#include <QtWidgets/QStatusBar>
#include <QtWidgets/QToolBar>
#include <QtWidgets/QWidget>

QT_BEGIN_NAMESPACE

class Ui_excelOperationDemoClass
{
public:
    QMenuBar *menuBar;
    QToolBar *mainToolBar;
    QWidget *centralWidget;
    QStatusBar *statusBar;

    void setupUi(QMainWindow *excelOperationDemoClass)
    {
        if (excelOperationDemoClass->objectName().isEmpty())
            excelOperationDemoClass->setObjectName(QString::fromUtf8("excelOperationDemoClass"));
        excelOperationDemoClass->resize(600, 400);
        menuBar = new QMenuBar(excelOperationDemoClass);
        menuBar->setObjectName(QString::fromUtf8("menuBar"));
        excelOperationDemoClass->setMenuBar(menuBar);
        mainToolBar = new QToolBar(excelOperationDemoClass);
        mainToolBar->setObjectName(QString::fromUtf8("mainToolBar"));
        excelOperationDemoClass->addToolBar(mainToolBar);
        centralWidget = new QWidget(excelOperationDemoClass);
        centralWidget->setObjectName(QString::fromUtf8("centralWidget"));
        excelOperationDemoClass->setCentralWidget(centralWidget);
        statusBar = new QStatusBar(excelOperationDemoClass);
        statusBar->setObjectName(QString::fromUtf8("statusBar"));
        excelOperationDemoClass->setStatusBar(statusBar);

        retranslateUi(excelOperationDemoClass);

        QMetaObject::connectSlotsByName(excelOperationDemoClass);
    } // setupUi

    void retranslateUi(QMainWindow *excelOperationDemoClass)
    {
        excelOperationDemoClass->setWindowTitle(QCoreApplication::translate("excelOperationDemoClass", "excelOperationDemo", nullptr));
    } // retranslateUi

};

namespace Ui {
    class excelOperationDemoClass: public Ui_excelOperationDemoClass {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_EXCELOPERATIONDEMO_H
