#ifndef TESTWINDOW_H
#define TESTWINDOW_H

#include <QMainWindow>
#include <QAbstractButton>
#include <QVariant>

class SExcel;

namespace Ui {
class TestWindow;
}

class TestWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit TestWindow(QWidget *parent = 0);
    ~TestWindow();

private slots:
    void on_pushButton_debugTest_clicked();
    void on_pushButton_openWorkDirectory_clicked();

    void on_pushButton_ExcelApp_Execute_clicked();
    void on_pushButton_ExcelApp_Quit_clicked();
    void on_pushButton_ExcelApp_SetVisible_clicked();

    void on_pushButton_WorkBooks_OpenFile_clicked();
    void on_pushButton_WorkBooks_NewWorkBook_clicked();
    void on_pushButton_WorkBooks_CloseAll_clicked();



    void on_pushButton_Range_SetProperty_clicked();

    void on_pushButton_Rows_SetProperty_clicked();

    void on_pushButton_Columns_SetProperty_clicked();

    void on_pushButton_Cell_SetProperty_clicked();

    void onValueTypeButtonClicked(QAbstractButton *button);
    QVariant getPropertyValue();

    void on_pushButton_Range_GetProperty_clicked();

    void on_pushButton_Rows_GetProperty_clicked();

    void on_pushButton_Columns_GetProperty_clicked();

    void on_pushButton_Cell_GetProperty_clicked();

    void on_pushButton_setActiveWorkBook_clicked();

    void on_pushButton_setActiveWorkSheet_clicked();

    void on_pushButton_newWorkSheet_clicked();

private:
    Ui::TestWindow *ui;

    SExcel *mExcel;

};

#endif // TESTWINDOW_H
