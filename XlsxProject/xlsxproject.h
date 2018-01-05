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

private slots:
	void readNames();
};

#endif // XLSXPROJECT_H
