#ifndef EXPORTARRANGE_H
#define EXPORTARRANGE_H

#include <QObject>
#include "xlsxdocument.h"

class ExportArrange : public QObject
{
	Q_OBJECT

public:
	ExportArrange(const QString &pExportPath, QObject *parent);
	~ExportArrange();

private:
	QXlsx::Document expoetXlsx;
	const QString &exportPath;

	void exportDataTime();
	void writeDateCell(const QString &pDate, char writeCell, const QXlsx::Format &dateStyle);
};

#endif // EXPORTARRANGE_H
