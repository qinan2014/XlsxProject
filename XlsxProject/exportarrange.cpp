#include "exportarrange.h"
#include <QDate>
#include "xlsxformat.h"

ExportArrange::ExportArrange(const QString &pExportPath, std::list<Single_Member> &pallMembers, QObject *parent)
	: QObject(parent), exportPath(pExportPath), allMembers(pallMembers)
{
	currentNameIndex = allMembers.begin();
	exportDataTime();
	arrangeOneMonth();
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
	if (pPos != ARRANGE_ERROR)
		return;
	Member_Arrange *tmpMember = searchArrangeDetail(pMember.s_name);
	if (tmpMember == NULL)
	{
		allArrangeMembers.push_back(Member_Arrange());
		tmpMember = (Member_Arrange *)&allArrangeMembers.back();
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
		Member_Arrange &tmpMember = (Member_Arrange &)it;
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
	
	Single_Member &tmpMember = (Single_Member &)currentNameIndex;
	++currentNameIndex;
	return &tmpMember;
}