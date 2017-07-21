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

    QButtonGroup *group = new QButtonGroup(this);
    group->addButton(ui->radioButton_text);
    group->addButton(ui->radioButton_integer);
    group->addButton(ui->radioButton_bool);
    connect(group, SIGNAL(buttonClicked(QAbstractButton*)), this, SLOT(onValueTypeButtonClicked(QAbstractButton*)));

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

void TestWindow::on_pushButton_Range_SetProperty_clicked()
{
    mExcel->setRangeProperty(ui->lineEdit_rangeName->text(), ui->lineEdit_propertyName->text(), getPropertyValue());
}

void TestWindow::on_pushButton_Rows_SetProperty_clicked()
{
    mExcel->setRowsProperty(ui->lineEdit_rangeName->text(), ui->lineEdit_propertyName->text(), getPropertyValue());
}

void TestWindow::on_pushButton_Columns_SetProperty_clicked()
{
    mExcel->setColumnsProperty(ui->lineEdit_rangeName->text(), ui->lineEdit_propertyName->text(), getPropertyValue());
}

void TestWindow::on_pushButton_Cell_SetProperty_clicked()
{
    mExcel->setCellProperty(ui->spinBox_cellRow->value(), ui->spinBox_cellColumn->value(),
                            ui->lineEdit_propertyName->text(), getPropertyValue());
}

void TestWindow::onValueTypeButtonClicked(QAbstractButton *button)
{
    if (button == ui->radioButton_bool) {
        ui->comboBox_propertyValue->setEnabled(true);
        ui->lineEdit_propertyValue->setEnabled(false);
        ui->spinBox_properyValue->setEnabled(false);
    } else if (button == ui->radioButton_integer) {
        ui->comboBox_propertyValue->setEnabled(false);
        ui->lineEdit_propertyValue->setEnabled(false);
        ui->spinBox_properyValue->setEnabled(true);
    } else if (button == ui->radioButton_text) {
        ui->comboBox_propertyValue->setEnabled(false);
        ui->lineEdit_propertyValue->setEnabled(true);
        ui->spinBox_properyValue->setEnabled(false);
    }
}

QVariant TestWindow::getPropertyValue()
{
    QVariant value;
    if (ui->radioButton_bool->isChecked()) {
        if (1 == ui->comboBox_propertyValue->currentIndex())
            value = true;
        else
            value = false;
    } else if (ui->radioButton_integer->isChecked()) {
        value = ui->spinBox_properyValue->value();
    } else if (ui->radioButton_text->isChecked()) {
        value = ui->lineEdit_propertyValue->text();
    }

    return value;
}

void TestWindow::on_pushButton_Range_GetProperty_clicked()
{
    QVariant value = mExcel->getRangeProperty(ui->lineEdit_rangeName->text(), ui->lineEdit_propertyName->text());
    qDebug() << value;
}

void TestWindow::on_pushButton_Rows_GetProperty_clicked()
{
    QVariant value = mExcel->getRowsProperty(ui->lineEdit_rangeName->text(), ui->lineEdit_propertyName->text());
    qDebug() << value;
}

void TestWindow::on_pushButton_Columns_GetProperty_clicked()
{
    QVariant value = mExcel->getColumnsProperty(ui->lineEdit_rangeName->text(), ui->lineEdit_propertyName->text());
    qDebug() << value;
}

void TestWindow::on_pushButton_Cell_GetProperty_clicked()
{
    QVariant value = mExcel->getCellProperty(ui->spinBox_cellRow->value(), ui->spinBox_cellColumn->value(), ui->lineEdit_propertyName->text());
    qDebug() << value;
}

void TestWindow::on_pushButton_setActiveWorkBook_clicked()
{
    mExcel->setActiveWorkBook(ui->lineEdit_index->text().toInt());
}

void TestWindow::on_pushButton_setActiveWorkSheet_clicked()
{
    mExcel->setActiveWorkSheet(ui->lineEdit_index->text().toInt());
}

void TestWindow::on_pushButton_newWorkSheet_clicked()
{
    mExcel->newWorkSheet();
}
