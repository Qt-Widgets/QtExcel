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



private:
    Ui::TestWindow *ui;

    SExcel *mExcel;

};

#endif // TESTWINDOW_H
