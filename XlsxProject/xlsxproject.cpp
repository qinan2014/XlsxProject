#include "xlsxproject.h"
#include <QFileDialog>
#include <QSettings>

XlsxProject::XlsxProject(QWidget *parent)
	: QWidget(parent)
{
	ui.setupUi(this);
	connect(ui.pbt_read_names, SIGNAL(clicked()), this, SLOT(readNames()));
}

XlsxProject::~XlsxProject()
{

}

void XlsxProject::readNames()
{
#define NOTHCHARCHPATH "noth_charch_path"
	QSettings settings("bbqinan", "NothCharch");
	QString oriPath = settings.value(NOTHCHARCHPATH).toString();
	QString fileName = QFileDialog::getOpenFileName(this, tr("Open File"), oriPath,
		tr("XLSX (*.xlsx)"));
	settings.setValue(NOTHCHARCHPATH, fileName);
#undef NOTHCHARCHPATH

}
