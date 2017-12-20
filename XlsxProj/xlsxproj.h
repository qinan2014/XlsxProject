#ifndef XLSXPROJ_H
#define XLSXPROJ_H

#include <QtWidgets/QWidget>
#include "ui_xlsxproj.h"

class XlsxProj : public QWidget
{
	Q_OBJECT

public:
	XlsxProj(QWidget *parent = 0);
	~XlsxProj();

private:
	Ui::XlsxProjClass ui;
};

#endif // XLSXPROJ_H
