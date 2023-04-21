#include "excelOperationDemo.h"
#include <QDir>
#include <QDebug>
#include <QAxObject>
#include <QDateTime>

excelOperationDemo::excelOperationDemo(QWidget *parent)
    : QMainWindow(parent)
{
    ui.setupUi(this);

	QString strFileDir = qApp->applicationDirPath() + QStringLiteral("/excelDir");
	QDir dir(strFileDir);
	if (!dir.exists())
	{
		bool resultMkdir = dir.mkdir(strFileDir);
		if (resultMkdir == false)
		{
			qDebug() << "mkdir excel directory is false";
		}
	}

	QString completePath = QString("%1/%2.xls").arg(strFileDir).arg(QDateTime::currentDateTime().toString("yyyyMMdd_hhmmss"));

	bool resultOperation = excelOperatorFunc(completePath);
	if (resultOperation == false)
	{
		qDebug() << "excel operation is false";
	}
}

excelOperationDemo::~excelOperationDemo()
{}


bool excelOperationDemo::excelOperatorFunc(QString filePath)
{
	HRESULT result = OleInitialize(nullptr); //头文件ole2.h
	OleUninitialize();
	QAxObject* excel = new QAxObject;
	if (excel->setControl("Excel.Application")) //连接Excel控件
	{
		excel->dynamicCall("SetVisible (bool Visible)", "false");//不显示窗体
		excel->setProperty("DisplayAlerts", false);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
		QAxObject* workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合

		workbooks->dynamicCall("Add");//新建一个工作簿

		QAxObject* workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
		QAxObject* worksheet = workbook->querySubObject("Worksheets(int)", 1);

		worksheet->setProperty("Name", QString::fromLocal8Bit("瑕疵信息"));

		QAxObject* cell;
		for (int nCol = 1; nCol <= 6; nCol++)
		{
			QStringList list;
			list << QString::fromLocal8Bit("序号") << QString::fromLocal8Bit("瑕疵名") << QString::fromLocal8Bit("长度/mm") << QString::fromLocal8Bit("宽度/mm")
				<< QString::fromLocal8Bit("x坐标/mm") << QString::fromLocal8Bit("y坐标/mm");

			cell = worksheet->querySubObject("Cells(int, int)", 1, nCol);
			cell->querySubObject("Font")->setProperty("Size", 12);
			cell->dynamicCall("SetValue(const QString&)", list[nCol - 1]);
			worksheet->querySubObject("Range(const QString&)", "1:1")->setProperty("RowHeight", 20);
		}

		for (int nRow = 1; nRow <= 3; nRow++)
		{
			QStringList list;
			list << QString::fromLocal8Bit("布匹长度/mm") << QString::fromLocal8Bit("布匹开始时间") << QString::fromLocal8Bit("布匹结束时间");
			cell = worksheet->querySubObject("Cells(int, int)", nRow, 9);
			cell->querySubObject("Font")->setProperty("Size", 12);
			cell->dynamicCall("SetValue(const QString&)", list[nRow - 1]);
			worksheet->querySubObject("Range(const QString&)", "1:1")->setProperty("RowHeight", 20);
			cell->setProperty("ColumnWidth", 15);
		}

		QAxObject* sheets = workbook->querySubObject("Sheets");
		int intCount = sheets->property("Count").toInt();
		QAxObject* lastSheet = sheets->querySubObject("Item(int)", intCount);
		sheets->dynamicCall("Add(QVariant)", lastSheet->asVariant());
		QAxObject* newSheet = sheets->querySubObject("Item(int)", intCount);
		lastSheet->dynamicCall("Move(QVariant)", newSheet->asVariant());

		newSheet->setProperty("Name", QString::fromLocal8Bit("得分"));
		for (int nCol = 1; nCol <= 3; nCol++)
		{
			QStringList list;
			list << QString::fromLocal8Bit("码数") << QString::fromLocal8Bit("瑕疵数量") << QString::fromLocal8Bit("扣分");

			cell = newSheet->querySubObject("Cells(int, int)", 1, nCol);
			cell->querySubObject("Font")->setProperty("Size", 12);
			cell->dynamicCall("SetValue(const QString&)", list[nCol - 1]);
			newSheet->querySubObject("Range(const QString&)", "1:1")->setProperty("RowHeight", 20);
		}

		workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(filePath));//保存至fileName
		workbook->dynamicCall("Close()");//关闭工作簿
		excel->dynamicCall("Quit()");
		delete excel;
		return true;
	}
    return false;
}