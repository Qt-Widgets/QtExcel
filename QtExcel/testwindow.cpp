#include "testwindow.h"
#include "ui_testwindow.h"
#include <Windows.h>
#include <QDebug>
#include <QUrl>
#include <QDesktopServices>
#include <QDir>
#include "sexcel.h"
#include <QFileDialog>


bool gWrapText = true;
bool gMergeCells = true;

bool gIsExcelAppVisible = true;


TestWindow::TestWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::TestWindow)
{
    ui->setupUi(this);

    mExcel = new SExcel(this);
}

TestWindow::~TestWindow()
{
    delete ui;
}

void TestWindow::on_pushButton_debugTest_clicked()
{
    qDebug() << "WorkBooks.Count = " << mExcel->workBooksCount();
    qDebug() << "WorkSheets.Count = " << mExcel->workSheetsCount();
}

void TestWindow::on_pushButton_openWorkDirectory_clicked()
{
    QDesktopServices::openUrl(QUrl::fromLocalFile(QDir::currentPath()));
}


void TestWindow::on_pushButton_ExcelApp_Execute_clicked()
{
    mExcel->execute();
}

void TestWindow::on_pushButton_ExcelApp_Quit_clicked()
{
    mExcel->quit();
}

void TestWindow::on_pushButton_ExcelApp_SetVisible_clicked()
{
    gIsExcelAppVisible = !gIsExcelAppVisible;
    mExcel->setVisible(gIsExcelAppVisible);
}

void TestWindow::on_pushButton_WorkBooks_OpenFile_clicked()
{
    QString filePath = QFileDialog::getOpenFileName(this);
    if (filePath.isEmpty())
        return;

    mExcel->open(filePath);
}

void TestWindow::on_pushButton_WorkBooks_NewWorkBook_clicked()
{
    mExcel->newWorkBook();
}

void TestWindow::on_pushButton_WorkBooks_CloseAll_clicked()
{
    mExcel->closeWorkBooks();
}

void TestWindow::on_pushButton_Range_SetHAlign_clicked()
{
    mExcel->setRangeHAlignment(1, 1, 3, 3, SExcel::HAlignCenter);
}

void TestWindow::on_pushButton_Range_SetVAlign_clicked()
{
    mExcel->setRangeVAlignment(1, 1, 3, 3, SExcel::VAlignCenter);
}

void TestWindow::on_pushButton_Range_SetMergeCells_clicked()
{
    gMergeCells = !gMergeCells;
    mExcel->setRangeMergeCells(1, 1, 3, 3, gMergeCells);
}

void TestWindow::on_pushButton_Range_SetWrapText_clicked()
{
    gWrapText = !gWrapText;
    mExcel->setRangeWrapText(1, 1, 3, 3, gWrapText);
}

void TestWindow::on_pushButton_Rows_SetHeight_clicked()
{
    mExcel->setRowsHeight(1, 3, 20);
}

void TestWindow::on_pushButton_Rows_SetHAlign_clicked()
{
    mExcel->setRowsHAlignment(1, 3, SExcel::HAlignLeft);
}

void TestWindow::on_pushButton_Rows_SetVAlign_clicked()
{
    mExcel->setRowsVAlignment(1, 3, SExcel::VAlignCenter);
}

void TestWindow::on_pushButton_Rows_SetMergeCells_clicked()
{
    gMergeCells = !gMergeCells;
    mExcel->setRowsMergeCells(1, 3, gMergeCells);
}

void TestWindow::on_pushButton_Rows_SetWrapText_clicked()
{
    gWrapText = !gWrapText;
    mExcel->setRowsWrapText(1, 3, gWrapText);
}

void TestWindow::on_pushButton_setColumnWidth_clicked()
{
    mExcel->setColumnsWidth(1, 3, 20);
}

void TestWindow::on_pushButton_Columns_SetHAlign_clicked()
{
    mExcel->setColumnsHAlignment(1, 3, SExcel::HAlignRight);
}

void TestWindow::on_pushButton_Columns_SetVAglin_clicked()
{
    mExcel->setColumnsVAlignment(1, 3, SExcel::VAlignBottom);
}

void TestWindow::on_pushButton_Columns_SetMergeCells_clicked()
{
    gMergeCells = !gMergeCells;
    mExcel->setColumnsMergeCells(1, 3, gMergeCells);
}

void TestWindow::on_pushButton_Columns_SetWrapText_clicked()
{
    gWrapText = !gWrapText;
    mExcel->setColumnsWrapText(1, 3, gWrapText);
}


