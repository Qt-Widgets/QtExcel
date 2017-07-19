#include "testwindow.h"
#include "ui_testwindow.h"
#include <Windows.h>
#include <QDebug>
#include <QUrl>
#include <QDesktopServices>
#include <QDir>
#include "sexcel.h"
#include <QFileDialog>

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

void TestWindow::on_pushButton_openWorkDirectory_clicked()
{
    QDesktopServices::openUrl(QUrl::fromLocalFile(QDir::currentPath()));
}

void TestWindow::on_pushButton_openExcelApp_clicked()
{
    mExcel->executeApp();
}

void TestWindow::on_pushButton_newWorkBook_clicked()
{
    mExcel->newWorkBook();
}

void TestWindow::on_pushButton_openExcelFile_clicked()
{
    QString path = QFileDialog::getOpenFileName(this);
    if (path.isEmpty())
        return;

    mExcel->open(path);
}

void TestWindow::on_pushButton_closeExcelApp_clicked()
{
    mExcel->quitApp();
}

void TestWindow::on_pushButton_closeWorkBook_clicked()
{
    mExcel->closeWorkBooks();
}

void TestWindow::on_pushButton_setRowHeight_clicked()
{
    mExcel->setRowsHeight(1, 5, 30);
}

void TestWindow::on_pushButton_setColumnWidth_clicked()
{
    mExcel->setColumnsWidth(1, 5, 30);
}

void TestWindow::on_pushButton_setRangeHAlign_clicked()
{
    mExcel->setRangeHAlignment(1, 1, 5, 5, SExcel::HAlignCenter);
}

void TestWindow::on_pushButton_setRangeVAlign_clicked()
{
    mExcel->setRangeVAlignment(1, 1, 5, 5, SExcel::VAlignCenter);
}

void TestWindow::on_pushButton_setRangeMergeCells_clicked()
{
    mExcel->setRangeMergeCells(1, 1, 5, 5, true);
}

void TestWindow::on_pushButton_setRangeWrapText_clicked()
{
    mExcel->setRangeWrapText(1, 1, 1, 1, true);
}

void TestWindow::on_pushButton_setRowsHAlign_clicked()
{
    mExcel->setRowsHAlignment(1, 1, SExcel::HAlignCenter);
}
