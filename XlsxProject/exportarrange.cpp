#include "exportarrange.h"
#include <QDate>
#include "xlsxformat.h"

ExportArrange::ExportArrange(const QString &pExportPath, QObject *parent)
	: QObject(parent), exportPath(pExportPath)
{
	exportDataTime();
}

ExportArrange::~ExportArrange()
{

}

void ExportArrange::exportDataTime()
{
	expoetXlsx.addSheet(QString::fromLocal8Bit("服务安排"));
	expoetXlsx.currentWorksheet()->setGridLinesVisible(true);

	QXlsx::Format dateHeaderStyle;
	dateHeaderStyle.setFontSize(14);
	dateHeaderStyle.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
	dateHeaderStyle.setVerticalAlignment(QXlsx::Format::AlignVCenter);
	expoetXlsx.setRowHeight(1, 18);
	
	QDate mBeginDate(2018, 1, 7);
	QDate mEndDate(2018, 2, 1);
	char beginColumCell = 'B';
	while (mEndDate > mBeginDate)
	{
		QString dataStr = mBeginDate.toString("yyyy.MM.dd");
		QString dataStr1 = dataStr + QString::fromLocal8Bit(" 4点");
		writeDateCell(dataStr1, beginColumCell, dateHeaderStyle);

		QString dataStr2 = dataStr + QString::fromLocal8Bit(" 6点");
		writeDateCell(dataStr2, beginColumCell + 2, dateHeaderStyle);

		mBeginDate = mBeginDate.addDays(7);
		beginColumCell += 4;
	}

	expoetXlsx.saveAs(exportPath);
}

void ExportArrange::writeDateCell(const QString &pDate, char writeCell, const QXlsx::Format &dateStyle)
{
	QString beginCell = QString("%1%2").arg(writeCell).arg('1');
	expoetXlsx.write(beginCell , pDate);
	char endCell = writeCell + 1;
	QString mergeEnd = QString("%1%2").arg(endCell).arg('1');
	QString mergeCell = beginCell + ':' + mergeEnd;
	expoetXlsx.mergeCells(mergeCell, dateStyle);
	expoetXlsx.setColumnWidth(writeCell - 64, writeCell - 63, 12);
	expoetXlsx.write(QString("%1%2").arg(writeCell).arg('2') , QString::fromLocal8Bit("预先安排"));
	expoetXlsx.write(QString("%1%2").arg(endCell).arg('2') , QString::fromLocal8Bit("实际服务"));
}