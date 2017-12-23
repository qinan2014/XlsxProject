#ifndef XLSXPROJECT_H
#define XLSXPROJECT_H

#include <QtWidgets/QWidget>
#include "ui_xlsxproject.h"

class XlsxProject : public QWidget
{
	Q_OBJECT

public:
	XlsxProject(QWidget *parent = 0);
	~XlsxProject();

private:
	Ui::XlsxProjectClass ui;
};

#endif // XLSXPROJECT_H
