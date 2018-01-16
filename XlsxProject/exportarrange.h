#ifndef EXPORTARRANGE_H
#define EXPORTARRANGE_H

#include <QObject>
#include "xlsxdocument.h"
#include "SingleMember.h"

class ExportArrange : public QObject
{
	Q_OBJECT

public:
	ExportArrange(const QString &pExportPath, std::list<Single_Member> &pallMembers, QObject *parent);
	~ExportArrange();

private:
	QXlsx::Document expoetXlsx;
	const QString &exportPath;
	std::list<Single_Member> &allMembers;
	std::list<Member_Arrange> allArrangeMembers;
	std::list<Single_Member>::iterator currentNameIndex; // 所有成员的迭代器


	void exportDataTime();
	void writeDateCell(int colum, const QString &pDate, const QXlsx::Format &dateStyle);
	void exportMemberArrange();
	void exportOneMember(int writeRow, const Member_Arrange &pMember);

	void arrangeOneMonth();
	void arrangeOneWeekMass(int pWeek);
	void insertArrange(int pWeek, int pMass, const Single_Member &pMember, int pPos);
	void arrangeOneMass(int pWeek, int pMass, bool charchPos[4], bool parkPos[2]);
	Single_Member *getOneSingleMemberByIterator();
	int arrangeOneMassCharch(int pWeek, int pMass, bool charchPos[4], Single_Member *inMember); // 返回已经安排的位置下标
	int arrangeOneMassPark(int pWeek, int pMass, bool parkPos[2], Single_Member *inMember);  // 返回已经安排的位置下标

	Member_Arrange *searchArrangeDetail(const std::wstring &arrangeName);
};

#endif // EXPORTARRANGE_H
