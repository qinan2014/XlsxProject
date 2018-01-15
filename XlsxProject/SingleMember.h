#ifndef SINGLE_MEMBER_H
#define SINGLE_MEMBER_H
#include <string.h>

#define WEEKSCOUNT 4
#define WEEK_ALL_SERVICE_TIME 2

#define ARRANGE_PARK_FLAG 0
#define ARRANGE_CHARCH_FLAG 1
#define ARRANGE_ERROR -1

struct Single_Member
{
	std::wstring s_name;
	short park_skill;
	bool is_lead;
	bool no_service_week[WEEKSCOUNT];
	
	Single_Member(){
		memset(no_service_week, 0, sizeof(no_service_week));
		park_skill = 0;
		is_lead = false;
	}
};

struct Member_Arrange 
{
	std::wstring s_name;
	bool pre_service_week[WEEKSCOUNT][WEEK_ALL_SERVICE_TIME];
	bool is_in_park[WEEKSCOUNT][WEEK_ALL_SERVICE_TIME];

	Member_Arrange()
	{
		memset(pre_service_week, 0, sizeof(pre_service_week));
		memset(is_in_park, 0, sizeof(is_in_park));
	}
};
#endif