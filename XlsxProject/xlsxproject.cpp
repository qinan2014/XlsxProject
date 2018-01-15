#include "xlsxproject.h"
#include <QFileDialog>
#include <QSettings>
#include "exportarrange.h"

XlsxProject::XlsxProject(QWidget *parent)
	: QWidget(parent)
{
	ui.setupUi(this);
	connect(ui.pbt_read_names, SIGNAL(clicked()), this, SLOT(readNames()));
	connect(ui.pbt_export, SIGNAL(clicked()), this, SLOT(exportList()));
}

XlsxProject::~XlsxProject()
{

}

void XlsxProject::readNames()
{
#define NOTHCHARCHPATH "noth_charch_path"
	QSettings settings("bbqinan", "NothCharchImport");
	QString oriPath = settings.value(NOTHCHARCHPATH).toString();
	QString fileName = QFileDialog::getOpenFileName(this, tr("Open File"), oriPath,
		tr("XLSX (*.xlsx)"));
	settings.setValue(NOTHCHARCHPATH, fileName);
#undef NOTHCHARCHPATH
	QXlsx::Document xlsx(fileName);
	std::list<Single_Member> members;
	readNameFromxls(xlsx, members);
	
}

void XlsxProject::readNameFromxls(QXlsx::Document &pDocXls, std::list<Single_Member> &pOutMembers)
{
	char tmpbuf[10];
	int nameIndex = 0;
	while (true)
	{
		sprintf(tmpbuf, "A%d", ++nameIndex);
		QString name = pDocXls.read(tmpbuf).toString();
		if (name.isEmpty())
			break;
		else {
			pOutMembers.push_back(Single_Member());
			Single_Member &tmpMember = pOutMembers.back();
			tmpMember.s_name = name.toStdWString();
		}
	}
}

void XlsxProject::exportList()
{
#define NOTHCHARCHPATH "noth_charch_path"
	QSettings settings("bbqinan", "NothCharchExport");
	QString oriPath = settings.value(NOTHCHARCHPATH).toString();
	QString fileName = QFileDialog::getSaveFileName(this, tr("Save File"), oriPath, tr("XLSX (*.xlsx)"));
	settings.setValue(NOTHCHARCHPATH, fileName);
#undef NOTHCHARCHPATH
	std::list<Single_Member> members;
	ExportArrange arrange(fileName, members, this);
}