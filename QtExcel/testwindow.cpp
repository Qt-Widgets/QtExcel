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
    mExcel->setActiveWorkSheet(2);
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


