#include "sexcel.h"
#include <Windows.h>
#include <QDebug>
#include <QDir>

SExcel::SExcel(QObject *parent) :
    QObject(parent)
{
    mExcelApp = NULL;
    mWorkBooks = NULL;
    mActiveWorkBook = NULL;
    mActiveWorkSheet = NULL;
    mIsVisible = true;
    mIsExecuted = false;

    mIsValid = initCOM();
}

SExcel::~SExcel()
{
    if (mIsExecuted) {
        quit();
    }
}

void SExcel::saveAsXLS(const QString &filePath)
{
    if (!mIsExecuted)
        return;

    mActiveWorkBook->dynamicCall("SaveAs(const QString &, int, const QString &, const QString &, bool, bool)",
                                  QDir::toNativeSeparators(filePath), 56, "", "", false, false);
}

void SExcel::saveAsXLSX(const QString &filePath)
{
    if (!mIsExecuted)
        return;

    mActiveWorkBook->dynamicCall("SaveAs(const QString &, int, const QString &, const QString &, bool, bool)",
                                 QDir::toNativeSeparators(filePath), 51, "", "", false, false);
}

void SExcel::open(const QString &filePath)
{
    if (!mIsExecuted)
        return;

    mWorkBooks->dynamicCall("Open(const QString &)", filePath);
    mActiveWorkBook = mExcelApp->querySubObject("ActiveWorkBook");
    mActiveWorkSheet = mActiveWorkBook->querySubObject("WorkSheets(int)", 1);
}

void SExcel::save()
{
    if (!mIsExecuted)
        return;

    mActiveWorkBook->dynamicCall("Save()");
}

int SExcel::workSheetsCount()
{
    if (!mIsExecuted)
        return 0;

    QAxObject *workSheets = mActiveWorkBook->querySubObject("WorkSheets");
    return workSheets->property("Count").toInt();
}

bool SExcel::execute()
{
    if (mIsExecuted)
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
    mIsExecuted = true;

    return true;
}

void SExcel::quit()
{
    if (mIsExecuted) {
        mExcelApp->dynamicCall("Quit(void)");
        deleteAxObjects();
        mIsExecuted = false;
    }
}

int SExcel::workBooksCount()
{
    if (!mIsExecuted)
        return 0;

    if (mWorkBooks)
        return mWorkBooks->property("Count").toInt();
    else
        return 0;
}

void SExcel::newWorkBook()
{
    if (!mIsExecuted)
        return;

    mWorkBooks->dynamicCall("Add");
    mActiveWorkBook = mExcelApp->querySubObject("ActiveWorkBook");
    mActiveWorkSheet = mActiveWorkBook->querySubObject("WorkSheets(int)", 1);
}

void SExcel::closeWorkBooks()
{
    if (!mIsExecuted)
        return;

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

void SExcel::setRangeProperty(const int startRow, const int startColumn, const int endRow, const int endColumn, const QString &propertyName, const QVariant &propertyValue)
{
    QAxObject *obj = getRange(startRow, startColumn, endRow, endColumn);
    QByteArray arrayPropertyName = propertyName.toLocal8Bit();
    obj->setProperty(arrayPropertyName.constData(), propertyValue);
    delete obj;
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

void SExcel::setRangeHAlignment(const int startRow, const int startColumn, const int endRow, const int endColumn, SExcel::HAlignment align)
{
    setRangeProperty(startRow, startColumn, endRow, endColumn, "HorizontalAlignment", align);
}

void SExcel::setRangeVAlignment(const int startRow, const int startColumn, const int endRow, const int endColumn, SExcel::VAlignment align)
{
    setRangeProperty(startRow, startColumn, endRow, endColumn, "VerticalAlignment", align);
}

void SExcel::setRangeMergeCells(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b)
{
    setRangeProperty(startRow, startColumn, endRow, endColumn, "MergeCells", b);
}

void SExcel::setRangeWrapText(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b)
{
    setRangeProperty(startRow, startColumn, endRow, endColumn, "WrapText", b);
}

void SExcel::setRangeFontProperty(const int startRow, const int startColumn, const int endRow, const int endColumn, const QString &propertyName, const QVariant &propertyValue)
{
    QAxObject *range = getRange(startRow, startColumn, endRow, endColumn);
    QAxObject *font = range->querySubObject("Font");
    QByteArray arrayPropertyName = propertyName.toLocal8Bit();
    font->setProperty(arrayPropertyName.constData(), propertyValue);
    delete font;
    delete range;
}

void SExcel::setRangeFontBold(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b)
{
    setRangeFontProperty(startRow, startColumn, endRow, endColumn, "Bold", b);
}

void SExcel::setRangeFontUnderline(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b)
{
    setRangeFontProperty(startRow, startColumn, endRow, endColumn, "Underline", b);
}

void SExcel::setRangeFontSize(const int startRow, const int startColumn, const int endRow, const int endColumn, const int size)
{
    setRangeFontProperty(startRow, startColumn, endRow, endColumn, "Size", size);
}

void SExcel::setRangeFontStrikethrough(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b)
{
    setRangeFontProperty(startRow, startColumn, endRow, endColumn, "Strikethrough", b);
}

void SExcel::setRangeFontSuperscript(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b)
{
    setRangeFontProperty(startRow, startColumn, endRow, endColumn, "Superscript", b);
}

void SExcel::setRangeFontSubscript(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b)
{
    setRangeFontProperty(startRow, startColumn, endRow, endColumn, "Subscript", b);
}

void SExcel::setRowsProperty(const int startRow, const int endRow, const QString &propertyName, const QVariant &propertyValue)
{
    QAxObject *obj = getRows(startRow, endRow);
    QByteArray arrayPropertyName = propertyName.toLocal8Bit();
    obj->setProperty(arrayPropertyName.constData(), propertyValue);
    delete obj;
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

void SExcel::setColumnsProperty(const int startColumn, const int endColumn, const QString &propertyName, const QVariant &propertyValue)
{
    QAxObject *obj = getColumns(startColumn, endColumn);
    QByteArray arrayPropertyName = propertyName.toLocal8Bit();
    obj->setProperty(arrayPropertyName.constData(), propertyValue);
    delete obj;
}

void SExcel::setColumnsWidth(const int startColumn, const int endColumn, const float width)
{
    setColumnsProperty(startColumn, endColumn, "ColumnWidth", width);
}

void SExcel::setColumnsHAlignment(const int startColumn, const int endColumn, SExcel::HAlignment align)
{
    setColumnsProperty(startColumn, endColumn, "HorizontalAlignment", align);
}

void SExcel::setColumnsVAlignment(const int startColumn, const int endColumn, SExcel::VAlignment align)
{
    setColumnsProperty(startColumn, endColumn, "VerticalAlignment", align);
}

void SExcel::setColumnsMergeCells(const int startColumn, const int endColumn, bool b)
{
    setColumnsProperty(startColumn, endColumn, "MergeCells", b);
}

void SExcel::setColumnsWrapText(const int startColumn, const int endColumn, bool b)
{
    setColumnsProperty(startColumn, endColumn, "WrapText", b);
}

void SExcel::setRowsHeight(const int startRow, const int endRow, const float height)
{
    setRowsProperty(startRow, endRow, "RowHeight", height);
}

void SExcel::setRowsHAlignment(const int startRow, const int endRow, SExcel::HAlignment align)
{
    setRowsProperty(startRow, endRow, "HorizontalAlignment", align);
}

void SExcel::setRowsVAlignment(const int startRow, const int endRow, SExcel::VAlignment align)
{
    setRowsProperty(startRow, endRow, "VerticalAlignment", align);
}

void SExcel::setRowsMergeCells(const int startRow, const int endRow, bool b)
{
    setRowsProperty(startRow, endRow, "MergeCells", b);
}

void SExcel::setRowsWrapText(const int startRow, const int endRow, bool b)
{
    setRowsProperty(startRow, endRow, "WrapText", b);
}

QAxObject *SExcel::getCell(const int row, const int column)
{
    return mActiveWorkSheet->querySubObject("Cells(int,int)", row, column);
}

void SExcel::setCellProperty(const int row, const int column, const QString &propertyName, const QVariant &propertyValue)
{
    QAxObject *obj = getCell(row, column);
    QByteArray arrayPropertyName = propertyName.toLocal8Bit();
    obj->setProperty(arrayPropertyName.constData(), propertyValue);
    delete obj;
}

void SExcel::setCellText(const int row, const int column, const QString &text)
{
    setCellProperty(row, column, "Value", text);
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



