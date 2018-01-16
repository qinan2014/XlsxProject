#include "exportarrange.h"
#include <QDate>
#include "xlsxformat.h"
#include "xlsxconditionalformatting.h"

ExportArrange::ExportArrange(const QString &pExportPath, std::list<Single_Member> &pallMembers, QObject *parent)
	: QObject(parent), exportPath(pExportPath), allMembers(pallMembers)
{
	currentNameIndex = allMembers.begin();
	exportDataTime();
	arrangeOneMonth();
	exportMemberArrange();
	exportEveryWeekStatic();
	exportMonthStatic();
	expoetXlsx.saveAs(exportPath);
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
	int beginColumCell = 2;
	while (mEndDate > mBeginDate)
	{
		QString dataStr = mBeginDate.toString("yyyy.MM.dd");
		QString dataStr1 = dataStr + QString::fromLocal8Bit(" 4点");
		writeDateCell(beginColumCell, dataStr1, dateHeaderStyle);
		
		QString dataStr2 = dataStr + QString::fromLocal8Bit(" 6点");
		writeDateCell(beginColumCell + 2, dataStr2, dateHeaderStyle);

		mBeginDate = mBeginDate.addDays(7);
		beginColumCell += 4;
	}
}

void ExportArrange::writeDateCell(int colum, const QString &pDate, const QXlsx::Format &dateStyle)
{
	expoetXlsx.write(1, colum , pDate);
	QXlsx::CellRange mergeRange(1, colum, 1, colum + 1);
	expoetXlsx.mergeCells(mergeRange, dateStyle);
	expoetXlsx.setColumnWidth(colum, colum + 1, 12);
	expoetXlsx.write(2, colum , QString::fromLocal8Bit("预先安排"));
	expoetXlsx.write(2, colum + 1 , QString::fromLocal8Bit("实际服务"));
}

void ExportArrange::arrangeOneMonth()
{
	for (int i = 0; i < WEEKSCOUNT; ++i)
	{
		arrangeOneWeekMass(i);
	}
}


void ExportArrange::arrangeOneWeekMass(int pWeek)
{
	bool charchService[2][4];
	bool parkService[2][2];
	memset(charchService, 0, sizeof(charchService));
	memset(parkService, 0, sizeof(parkService));
	arrangeOneMass(pWeek, 0, charchService[0], parkService[0]);
	arrangeOneMass(pWeek, 1, charchService[1], parkService[1]);
}

void ExportArrange::arrangeOneMass(int pWeek, int pMass, bool charchPos[4], bool parkPos[2])
{
	Single_Member *tmpMember = getOneSingleMemberByIterator();
	int charchPosIndex = -1;
	int parkPosIndex = -1;
	while ((charchPosIndex < 3) || (parkPosIndex < 1))
	{
		if (parkPosIndex < 1 && tmpMember->park_skill > 1)
			parkPosIndex = arrangeOneMassPark(pWeek, pMass, parkPos, tmpMember);
		else
			charchPosIndex = arrangeOneMassCharch(pWeek, pMass, charchPos, tmpMember);
		tmpMember = getOneSingleMemberByIterator();
	}
}

int ExportArrange::arrangeOneMassCharch(int pWeek, int pMass, bool charchPos[4], Single_Member *inMember)
{
	for (int i = 0; i < 4; ++i)
	{
		if (!charchPos[i])
		{
			charchPos[i] = true;
			insertArrange(pWeek, pMass, *inMember, ARRANGE_CHARCH_FLAG);
			return i;
		}
	}
	return 3;
}

int ExportArrange::arrangeOneMassPark(int pWeek, int pMass, bool parkPos[2], Single_Member *inMember)
{
	for (int i = 0; i < 2; ++i)
	{
		if (!parkPos[i])
		{
			parkPos[i] = true;
			insertArrange(pWeek, pMass, *inMember, ARRANGE_PARK_FLAG);
			return i;
		}
	}
	return 1;
}

void ExportArrange::insertArrange(int pWeek, int pMass, const Single_Member &pMember, int pPos)
{
	if (pPos == ARRANGE_ERROR)
		return;
	Member_Arrange *tmpMember = searchArrangeDetail(pMember.s_name);
	if (tmpMember == NULL)
	{
		allArrangeMembers.push_back(Member_Arrange());
		Member_Arrange &arrangeMember = allArrangeMembers.back();
		tmpMember = &arrangeMember;
	}
	
	tmpMember->s_name = pMember.s_name;
	if (pPos == ARRANGE_PARK_FLAG)
		tmpMember->is_in_park[pWeek][pMass] = true;
	else 
		tmpMember->pre_service_week[pWeek][pMass] = true;
}

Member_Arrange *ExportArrange::searchArrangeDetail(const std::wstring &arrangeName)
{
	Member_Arrange *mArrange = NULL;
	for (std::list<Member_Arrange>::iterator it = allArrangeMembers.begin(); it != allArrangeMembers.end(); ++it)
	{
		Member_Arrange &tmpMember = (Member_Arrange &)*it;
		if (tmpMember.s_name == arrangeName)
		{
			mArrange = &tmpMember;
			break;
		}
	}
	return mArrange;
}

Single_Member *ExportArrange::getOneSingleMemberByIterator()
{
	if (currentNameIndex == allMembers.end())
		currentNameIndex = allMembers.begin();
	
	Single_Member &tmpMember = (Single_Member &)*currentNameIndex;
	++currentNameIndex;
	return &tmpMember;
}

void ExportArrange::exportMemberArrange()
{
	int beginRow = 3;
	for (std::list<Member_Arrange>::iterator it = allArrangeMembers.begin(); it != allArrangeMembers.end(); ++it)
	{
		Member_Arrange &tmpMember = (Member_Arrange &)*it;
		exportOneMember(beginRow, tmpMember);
		++beginRow;
	}
	QXlsx::DataValidation validation(QXlsx::DataValidation::List, QXlsx::DataValidation::Between, "33", "35");
	validation.setFormula1(QString::fromLocal8Bit("\"_,堂内,献宜,车场\""));
	int beginColumn = 2;
	for (int i = 0; i < WEEKSCOUNT; ++i)
	{
		for (int j = 0; j < WEEK_ALL_SERVICE_TIME; ++j)
		{
			QXlsx::CellRange rangePre(1, beginColumn, beginRow - 1, beginColumn);
			validation.addRange(rangePre);
			QXlsx::CellRange rangeAct(1, beginColumn + 1, beginRow - 1, beginColumn + 1);
			validation.addRange(rangeAct);
			beginColumn += 2;
		}
	}
	expoetXlsx.addDataValidation(validation);
}

void ExportArrange::exportOneMember(int writeRow, const Member_Arrange &pMember)
{
	QString tmpName = QString::fromStdWString(pMember.s_name);
	expoetXlsx.write(writeRow, 1 , tmpName);
	int beginColumn = 2;
	for (int i = 0; i < WEEKSCOUNT; ++i)
	{
		for (int j = 0; j < WEEK_ALL_SERVICE_TIME; ++j)
		{
			if (pMember.is_in_park[i][j])
				expoetXlsx.write(writeRow, beginColumn , QString::fromLocal8Bit("车场"));
			else if (pMember.pre_service_week[i][j])
				expoetXlsx.write(writeRow, beginColumn , QString::fromLocal8Bit("堂内"));
			beginColumn += 2;
		}
	}

}

void ExportArrange::exportEveryWeekStatic()
{
	int membersz = allMembers.size();
	int rowStatic = membersz + 2 + 1;
	QXlsx::Format format;
	format.setFontColor(QColor(Qt::blue));
	format.setFontSize(12);
	expoetXlsx.write(rowStatic, 1, QString::fromLocal8Bit("服务统计"));
	expoetXlsx.setRowFormat(rowStatic, rowStatic, format);

	int beginColumn = 2;
	for (int i = 0; i < WEEKSCOUNT; ++i)
	{
		for (int j = 0; j < WEEK_ALL_SERVICE_TIME; ++j)
		{
			char rangchar[20];
			sprintf(rangchar, "%c%d:%c%d", beginColumn + 64, 3, beginColumn + 64, rowStatic - 1);
			char selectedchar[100];
			sprintf(selectedchar, "=COUNTIF(%s,\"车场\")+COUNTIF(%s,\"献宜\")+COUNTIF(%s,\"堂内\")", rangchar, rangchar, rangchar);
			QString funcStr = QString::fromLocal8Bit(selectedchar);
			expoetXlsx.write(rowStatic, beginColumn, funcStr);
			++beginColumn;
			sprintf(rangchar, "%c%d:%c%d", beginColumn + 64, 3, beginColumn + 64, rowStatic - 1);
			sprintf(selectedchar, "=COUNTIF(%s,\"车场\")+COUNTIF(%s,\"献宜\")+COUNTIF(%s,\"堂内\")", rangchar, rangchar, rangchar);
			expoetXlsx.write(rowStatic, beginColumn, funcStr);
			++beginColumn;
		}
	}
}

void ExportArrange::exportMonthStatic()
{
	QXlsx::Format dateHeaderStyle;
	dateHeaderStyle.setFontSize(10);
	dateHeaderStyle.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
	dateHeaderStyle.setVerticalAlignment(QXlsx::Format::AlignVCenter);
	int beginStaticColumn = WEEKSCOUNT * 4 + 1 + 1;

	expoetXlsx.write(1, beginStaticColumn , QString::fromLocal8Bit("安排次数"));
	QXlsx::CellRange mergeRange(1, beginStaticColumn, 2, beginStaticColumn);
	expoetXlsx.mergeCells(mergeRange, dateHeaderStyle);
	
	expoetXlsx.write(1, beginStaticColumn + 1 , QString::fromLocal8Bit("安排车场次数"));
	QXlsx::CellRange mergeRange1(1, beginStaticColumn + 1, 2, beginStaticColumn + 1);
	expoetXlsx.mergeCells(mergeRange1, dateHeaderStyle);
}