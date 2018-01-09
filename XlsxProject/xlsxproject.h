#ifndef XLSXPROJECT_H
#define XLSXPROJECT_H

#include <QtWidgets/QWidget>
#include "ui_xlsxproject.h"
#include "xlsxdocument.h"
#include "SingleMember.h"

class XlsxProject : public QWidget
{
	Q_OBJECT

public:
	XlsxProject(QWidget *parent = 0);
	~XlsxProject();

private:
	Ui::XlsxProjectClass ui;

	void readNameFromxls(QXlsx::Document &pDocXls, std::list<Single_Member> &pOutMembers);

private slots:
	void readNames();
	void exportList();
};

#endif // XLSXPROJECT_H
