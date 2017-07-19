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
    void on_pushButton_openWorkDirectory_clicked();

    void on_pushButton_openExcelApp_clicked();

    void on_pushButton_newWorkBook_clicked();

    void on_pushButton_openExcelFile_clicked();

    void on_pushButton_closeExcelApp_clicked();

    void on_pushButton_closeWorkBook_clicked();

    void on_pushButton_setRowHeight_clicked();

    void on_pushButton_setColumnWidth_clicked();

    void on_pushButton_setRangeHAlign_clicked();

    void on_pushButton_setRangeVAlign_clicked();

    void on_pushButton_setRangeMergeCells_clicked();

    void on_pushButton_setRangeWrapText_clicked();

    void on_pushButton_setRowsHAlign_clicked();

private:
    Ui::TestWindow *ui;

    SExcel *mExcel;
};

#endif // TESTWINDOW_H
