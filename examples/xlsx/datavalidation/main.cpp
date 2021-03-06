﻿#include <QtCore>
#include "xlsxdocument.h"
#include "xlsxdatavalidation.h"

QTXLSX_USE_NAMESPACE

int main()
{
    Document xlsx;
    xlsx.write("A1", "A2 and A3:E5 only accept the number between 33 and 99");

    //![1]
    //DataValidation validation(DataValidation::Whole, DataValidation::Between, "33", "99");
    //validation.addRange("A2");
    //validation.addRange("A3:E5");
    //validation.setPromptMessage("Please Input Integer between 33 and 99");
    //xlsx.addDataValidation(validation);
    //![2]  下拉列表可以了
	DataValidation validation(DataValidation::List, DataValidation::Between, "33", "35");
	validation.setFormula1("\"item1,item2,item3\"");
	validation.addRange("A2");
	validation.addRange("A3:E5");
	validation.setPromptMessage("Please Input Integer between 33 and 99");
	xlsx.addDataValidation(validation);

	xlsx.write("A6", "=A1");

    xlsx.save();
    return 0;
}
