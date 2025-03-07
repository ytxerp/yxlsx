#include <QCoreApplication>
#include <QDir>

#include "document.h"

int main(int argc, char* argv[])
{
    QCoreApplication app(argc, argv);

    // [1]  Writing excel file(*.xlsx)
    qDebug() << "------------------[1]------------------------";

    YXlsx::Document test1 {};
    auto book1 { test1.GetWorkbook() };

    book1->GetCurrentWorksheet()->Write(1, 1, "Hello Qt!");
    book1->GetCurrentWorksheet()->Write(1, 2, 2);
    book1->GetCurrentWorksheet()->Write(1, 3, true);
    book1->GetCurrentWorksheet()->Write(1, 4, QDateTime::currentDateTime());
    book1->GetCurrentWorksheet()->Write("b1", 2);

    if (!test1.Save("Test1.xlsx")) {
        qDebug() << "Failed to write xlsx file";
    }

    // [2] Reading excel file(*.xlsx)
    qDebug() << "------------------[2]------------------------";

    YXlsx::Document test2("Test1.xlsx");
    auto book2 { test2.GetWorkbook() };

    if (test2.IsLoadXlsx()) {
        auto value { book2->GetCurrentWorksheet()->Read(1, 1) };
        qDebug() << "Cell(1,1) is " << value;

        value = book2->GetCurrentWorksheet()->Read(1, 2);
        qDebug() << "Cell(1,2) is " << value;

        value = book2->GetCurrentWorksheet()->Read(1, 3);
        qDebug() << "Cell(1,3) is " << value;

        value = book2->GetCurrentWorksheet()->Read(1, 4);
        qDebug() << "Cell(1,4) is " << value;

        value = book2->GetCurrentWorksheet()->Read("B1");
        qDebug() << "Cell(2,1) is " << value;

    } else {
        qDebug() << "Failed to load xlsx file.";
    }

    // [3]
    qDebug() << "------------------[3]------------------------";

    YXlsx::Document test3("Test2.xlsx");
    auto book3 { test3.GetWorkbook() };
    QList<int> list { 1, 2, 3 };
    book3->GetCurrentWorksheet()->WriteColumn(1, 1, list);
    book3->GetCurrentWorksheet()->WriteRow(1, 4, list);

    QList<QString> list3 { "hello", "world", "YTX" };
    QList<QVariant> list4 { 1, "world", 4.4, QVariant() };

    book3->GetCurrentWorksheet()->WriteRow(1, 8, list3);
    book3->GetCurrentWorksheet()->WriteRow(2, 8, list4);

    QVector<QVariant> list2 = { "h", "e", "l", "l", "o" };
    book3->GetCurrentWorksheet()->WriteColumn(1, 2, list2);
    book3->GetCurrentWorksheet()->WriteColumn(1, 3, list2);
    book3->GetCurrentWorksheet()->WriteColumn(1, 4, list2);

    book3->GetCurrentWorksheet()->WriteColumn(10, 10, list2);

    if (!test3.Save()) {
        qDebug() << "Failed to write xlsx file";
    }

    // [4]
    qDebug() << "------------------[4]------------------------";

    YXlsx::Document test4("Test3.xlsx");
    auto workbook4 { test4.GetWorkbook() };

    // current sheet is Sheet1(default sheet)
    for (int i = 1; i < 20; ++i) {
        for (int j = 1; j < 15; ++j) {
            workbook4->GetCurrentWorksheet()->Write(i, j, QString("R %1 C %2").arg(i).arg(j));
        }
    }

    workbook4->AppendSheet();
    workbook4->GetCurrentWorksheet()->Write(2, 2, "Hello Qt Xlsx");

    workbook4->AppendSheet();
    workbook4->GetCurrentWorksheet()->Write(3, 3, "This will be deleted...");

    workbook4->AppendSheet("HiddenSheet");
    workbook4->GetCurrentWorksheet()->Write(YXlsx::Coordinate("A1"), "This sheet is hidden.");

    workbook4->AppendSheet("VeryHiddenSheet");
    workbook4->GetCurrentWorksheet()->Write(YXlsx::Coordinate("A1"), "This sheet is very hidden.");

    workbook4->RenameSheet("HiddenSheet", "Hello World");

    if (!test4.Save()) {
        qDebug() << "Failed to write excel.";
    }

    qDebug() << "Sheet Names" << workbook4->GetSheetName();
    qDebug() << "Document Properties" << test4.GetProperty();
    qDebug() << "Set current sheet index: 2." << workbook4->SetCurrentSheet(2);
    qDebug() << "Index 2' name" << workbook4->GetCurrentSheet()->GetSheetName();

    qDebug() << "------------------[5]------------------------";

    return 0;
}
