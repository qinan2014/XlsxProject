#ifndef SINGLE_MEMBER_H
#define SINGLE_MEMBER_H
#include <string.h>

#define WEEKSCOUNT 4
#define WEEK_ALL_SERVICE_TIME 2
struct Single_Member
{
	std::wstring s_name;
	bool pre_service_week[WEEKSCOUNT][WEEK_ALL_SERVICE_TIME];
	bool is_in_park[WEEKSCOUNT][WEEK_ALL_SERVICE_TIME];
	short park_skill;
	bool is_lead;
	Single_Member(){
		memset(pre_service_week, 0, sizeof(pre_service_week));
		memset(is_in_park, 0, sizeof(is_in_park));
		park_skill = 0;
		is_lead = false;
	}
};
#endif