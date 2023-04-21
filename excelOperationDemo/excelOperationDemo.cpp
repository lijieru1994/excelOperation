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
	HRESULT result = OleInitialize(nullptr); //ͷ�ļ�ole2.h
	OleUninitialize();
	QAxObject* excel = new QAxObject;
	if (excel->setControl("Excel.Application")) //����Excel�ؼ�
	{
		excel->dynamicCall("SetVisible (bool Visible)", "false");//����ʾ����
		excel->setProperty("DisplayAlerts", false);//����ʾ�κξ�����Ϣ�����Ϊtrue��ô�ڹر��ǻ�������ơ��ļ����޸ģ��Ƿ񱣴桱����ʾ
		QAxObject* workbooks = excel->querySubObject("WorkBooks");//��ȡ����������

		workbooks->dynamicCall("Add");//�½�һ��������

		QAxObject* workbook = excel->querySubObject("ActiveWorkBook");//��ȡ��ǰ������
		QAxObject* worksheet = workbook->querySubObject("Worksheets(int)", 1);

		worksheet->setProperty("Name", QString::fromLocal8Bit("覴���Ϣ"));

		QAxObject* cell;
		for (int nCol = 1; nCol <= 6; nCol++)
		{
			QStringList list;
			list << QString::fromLocal8Bit("���") << QString::fromLocal8Bit("覴���") << QString::fromLocal8Bit("����/mm") << QString::fromLocal8Bit("���/mm")
				<< QString::fromLocal8Bit("x����/mm") << QString::fromLocal8Bit("y����/mm");

			cell = worksheet->querySubObject("Cells(int, int)", 1, nCol);
			cell->querySubObject("Font")->setProperty("Size", 12);
			cell->dynamicCall("SetValue(const QString&)", list[nCol - 1]);
			worksheet->querySubObject("Range(const QString&)", "1:1")->setProperty("RowHeight", 20);
		}

		for (int nRow = 1; nRow <= 3; nRow++)
		{
			QStringList list;
			list << QString::fromLocal8Bit("��ƥ����/mm") << QString::fromLocal8Bit("��ƥ��ʼʱ��") << QString::fromLocal8Bit("��ƥ����ʱ��");
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

		newSheet->setProperty("Name", QString::fromLocal8Bit("�÷�"));
		for (int nCol = 1; nCol <= 3; nCol++)
		{
			QStringList list;
			list << QString::fromLocal8Bit("����") << QString::fromLocal8Bit("覴�����") << QString::fromLocal8Bit("�۷�");

			cell = newSheet->querySubObject("Cells(int, int)", 1, nCol);
			cell->querySubObject("Font")->setProperty("Size", 12);
			cell->dynamicCall("SetValue(const QString&)", list[nCol - 1]);
			newSheet->querySubObject("Range(const QString&)", "1:1")->setProperty("RowHeight", 20);
		}

		workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(filePath));//������fileName
		workbook->dynamicCall("Close()");//�رչ�����
		excel->dynamicCall("Quit()");
		delete excel;
		return true;
	}
    return false;
}