#ifndef TESTWINDOW_H
#define TESTWINDOW_H

#include <QMainWindow>

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

    void on_pushButton_Range_SetHAlign_clicked();
    void on_pushButton_Range_SetVAlign_clicked();
    void on_pushButton_Range_SetMergeCells_clicked();
    void on_pushButton_Range_SetWrapText_clicked();

    void on_pushButton_Rows_SetHeight_clicked();
    void on_pushButton_Rows_SetHAlign_clicked();
    void on_pushButton_Rows_SetVAlign_clicked();
    void on_pushButton_Rows_SetMergeCells_clicked();
    void on_pushButton_Rows_SetWrapText_clicked();

    void on_pushButton_setColumnWidth_clicked();

    void on_pushButton_Columns_SetHAlign_clicked();
    void on_pushButton_Columns_SetVAglin_clicked();
    void on_pushButton_Columns_SetMergeCells_clicked();
    void on_pushButton_Columns_SetWrapText_clicked();



private:
    Ui::TestWindow *ui;

    SExcel *mExcel;

};

#endif // TESTWINDOW_H
