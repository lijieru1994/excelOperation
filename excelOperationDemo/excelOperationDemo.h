#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_excelOperationDemo.h"
#include <Ole2.h>

class excelOperationDemo : public QMainWindow
{
    Q_OBJECT

public:
    excelOperationDemo(QWidget *parent = nullptr);
    ~excelOperationDemo();

    bool excelOperatorFunc(QString filePath);
private:
    Ui::excelOperationDemoClass ui;
};
