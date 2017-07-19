#include "sexcel.h"
#include <Windows.h>
#include <QDebug>
#include <QDir>

SExcel::SExcel(QObject *parent) :
    QObject(parent)
{
    mExcelApp = NULL;
    mWorkBooks = NULL;
    mWorkBooksCount = 0;
    mActiveWorkBook = NULL;
    mActiveWorkSheet = NULL;
    mIsVisible = true;
    mIsAppExecuted = false;

    mIsValid = initCOM();
}

SExcel::~SExcel()
{
    if (mIsAppExecuted) {
        quitApp();
    }
}

void SExcel::saveAsXLS(const QString &filePath)
{
    mActiveWorkBook->dynamicCall("SaveAs(const QString &, int, const QString &, const QString &, bool, bool)",
                                  QDir::toNativeSeparators(filePath), 56, "", "", false, false);
}

void SExcel::saveAsXLSX(const QString &filePath)
{
    mActiveWorkBook->dynamicCall("SaveAs(const QString &, int, const QString &, const QString &, bool, bool)",
                                 QDir::toNativeSeparators(filePath), 51, "", "", false, false);
}

void SExcel::open(const QString &filePath)
{
    mWorkBooks->dynamicCall("Open(const QString &)", filePath);
    ++ mWorkBooksCount;

    mActiveWorkBook = mExcelApp->querySubObject("ActiveWorkBook");
    mActiveWorkSheet = mActiveWorkBook->querySubObject("WorkSheets(int)", 1);
}

void SExcel::save()
{
    mActiveWorkBook->dynamicCall("Save()");
}

bool SExcel::executeApp()
{
    if (mIsAppExecuted)
        return true;

    // 打开excel程序
    mExcelApp = new QAxObject("Excel.Application");
    if (mExcelApp->isNull()) {
        mErrorString = QString::fromLocal8Bit("无法打开Excel应用程序,请确认您的电脑已经正确安装了Excel");
        mIsValid = false;
        return false;
    }

    //
    mExcelApp->setProperty("Visible", mIsVisible);
    mWorkBooks = mExcelApp->querySubObject("WorkBooks");
    mIsAppExecuted = true;

    return true;
}

void SExcel::quitApp()
{
    if (mIsAppExecuted) {
        mExcelApp->dynamicCall("Quit(void)");
        deleteAxObjects();
        mIsAppExecuted = false;
    }
}

void SExcel::newWorkBook()
{
    mWorkBooks->dynamicCall("Add");
    mActiveWorkBook = mExcelApp->querySubObject("ActiveWorkBook");
    mActiveWorkSheet = mActiveWorkBook->querySubObject("WorkSheets(int)", 1);
}

void SExcel::closeWorkBooks()
{
    mWorkBooks->dynamicCall("Close");
}

QAxObject *SExcel::getRange(const QString &name)
{
    return mActiveWorkSheet->querySubObject("Range(const QString &)", name);
}

QAxObject *SExcel::getRange(const int startRow, const int startColumn,
            const int endRow, const int endColumn)
{
    QString startCellName = QString("%1%2")
            .arg((char)(startColumn - 1 + 'A'))
            .arg(startRow);
    QString endCellName = QString("%1%2")
            .arg((char)(endColumn - 1 + 'A'))
            .arg(endRow);

    QString rangeName;
    if (startCellName == endCellName)
        rangeName = startCellName;
    else
        rangeName = QString("%1:%2").arg(startCellName).arg(endCellName);

    return getRange(rangeName);
}

QAxObject *SExcel::getRows(const QString &name)
{
    return mActiveWorkSheet->querySubObject("Rows(const QString &)", name);
}

QAxObject *SExcel::getRows(const int startRow, const int endRow)
{
    return getRows(
                QString("%1:%2").arg(startRow).arg(endRow)
                );
}

QAxObject *SExcel::getColumns(const QString &name)
{
    return mActiveWorkSheet->querySubObject("Columns(const QString &)", name);
}

QAxObject *SExcel::getColumns(const int startColumn, const int endColumn)
{
    return getColumns(
                QString("%1:%2").arg((char)(startColumn - 1 + 'A')).arg((char)(endColumn - 1 + 'A'))
                );
}

void SExcel::setRowsHeight(const int startRow, const int endRow, const float height)
{
    QAxObject *obj = getRows(startRow, endRow);
    obj->setProperty("RowHeight", height);
    delete obj;
}

void SExcel::setRowsHAlignment(const int startRow, const int endRow, SExcel::HAlignment align)
{
    QAxObject *obj = getRows(startRow, endRow);
    obj->setProperty("HorizontalAlignment", align);
    delete obj;
}

void SExcel::setRowsVAlignment(const int startRow, const int endRow, SExcel::VAlignment align)
{
    QAxObject *obj = getRows(startRow, endRow);
    obj->setProperty("VerticalAlignment", align);
    delete obj;
}

void SExcel::setColumnsWidth(const int startColumn, const int endColumn, const float width)
{
    QAxObject *obj = getColumns(startColumn, endColumn);
    obj->setProperty("ColumnWidth", width);
    delete obj;
}

void SExcel::setColumnsHAlignment(const int startColumn, const int endColumn, SExcel::HAlignment align)
{
    QAxObject *obj = getColumns(startColumn, endColumn);
    obj->setProperty("HorizontalAlignment", align);
    delete obj;
}

void SExcel::setColumnsVAlignment(const int startColumn, const int endColumn, SExcel::VAlignment align)
{
    QAxObject *obj = getColumns(startColumn, endColumn);
    obj->setProperty("VerticalAlignment", align);
    delete obj;
}

QAxObject *SExcel::getCell(const int row, const int column)
{
    return mActiveWorkSheet->querySubObject("Cells(int,int)", row, column);
}

void SExcel::setCellText(const int row, const int column, const QString &text)
{
    QAxObject *cell = getCell(row, column);
    cell->dynamicCall("SetValue(const QString &)", text);
    delete cell;
}

void SExcel::setRangeHAlignment(const int startRow, const int startColumn, const int endRow, const int endColumn, SExcel::HAlignment align)
{
    QAxObject *obj = getRange(startRow, startColumn, endRow, endColumn);
    obj->setProperty("HorizontalAlignment", align);
    delete obj;
}

void SExcel::setRangeVAlignment(const int startRow, const int startColumn, const int endRow, const int endColumn, SExcel::VAlignment align)
{
    QAxObject *obj = getRange(startRow, startColumn, endRow, endColumn);
    obj->setProperty("VerticalAlignment", align);
    delete obj;
}

void SExcel::setRangeMergeCells(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b)
{
    QAxObject *obj = getRange(startRow, startColumn, endRow, endColumn);
    obj->setProperty("MergeCells", b);
    delete obj;
}

void SExcel::setRangeWrapText(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b)
{
    QAxObject *obj = getRange(startRow, startColumn, endRow, endColumn);
    obj->setProperty("WrapText", b);
    delete obj;
}

bool SExcel::initCOM()
{
    HRESULT r = OleInitialize(0);
    if (S_OK != r && S_FALSE != r) {
        mErrorString = QString::fromLocal8Bit("无法初始化Windows COM组件");
        return false;
    }

    return true;
}

void SExcel::deleteAxObjects()
{
    if (mExcelApp) {
        delete mExcelApp;
        mExcelApp = NULL;
    }
    // 其它QAxObject都是mExcelApp的后代,
    // 当mExcelApp释放时,它们都会被释放,
    // 如果重复释放,将会导致程序崩溃
}



