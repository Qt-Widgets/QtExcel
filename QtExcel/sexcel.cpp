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

void SExcel::setActiveWorkSheet(const int index)
{
    if (!mIsExecuted)
        return;

    int count = workSheetsCount();
    if (count <= 0) {
        return;
    }  else {
        if (index >= 1 && index <= count) {
            mActiveWorkBook->dynamicCall("Select(const QString &)", "Sheet2");
            mActiveWorkSheet = mExcelApp->querySubObject("ActiveWorkSheet");
        }
    }

}

int SExcel::workSheetsCount()
{
    if (!mIsExecuted)
        return 0;

    if (mActiveWorkBook) {
        QAxObject *workSheets = mActiveWorkBook->querySubObject("WorkSheets");
        int count = workSheets->property("Count").toInt();
        delete workSheets;
        return count;
    } else {
        return 0;
    }
}

void SExcel::setActiveWorkSheetName(const QString &name)
{
    if (!mIsExecuted)
        return;

    mActiveWorkSheet->setProperty("Name", name);
}

QString SExcel::getRangeName(const int startRow, const int startColumn, const int endRow, const int endColumn)
{
    QString rangeName;

    QString startCellName = QString("%1%2")
            .arg((char)(startColumn - 1 + 'A'))
            .arg(startRow);

    QString endCellName = QString("%1%2")
            .arg((char)(endColumn - 1 + 'A'))
            .arg(endRow);

    if (startCellName == endCellName)
        rangeName = startCellName;
    else
        rangeName = QString("%1:%2").arg(startCellName).arg(endCellName);

    return rangeName;
}

bool SExcel::execute()
{
    // 如果已经运行
    if (mIsExecuted)
        return true;

    // 打开excel程序
    mExcelApp = new QAxObject("Excel.Application");
    if (mExcelApp->isNull()) {
        mErrorString = QString::fromLocal8Bit("无法打开Excel应用程序,请确认您的电脑已经正确安装了Excel");
        mIsValid = false;
        return false;
    }

    // 设置app窗口可见性
    mExcelApp->setProperty("Visible", mIsVisible);

    // 得到所有工作簿对象
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

    // 创建的工作簿被激活,成为活动工作簿
    mActiveWorkBook = mExcelApp->querySubObject("ActiveWorkBook");

    // 活动工作簿的第一个工作表默认为活动工作表
    mActiveWorkSheet = mActiveWorkBook->querySubObject("ActiveSheet");
}

void SExcel::closeWorkBooks()
{
    if (!mIsExecuted)
        return;

    mWorkBooks->dynamicCall("Close");
}

QAxObject *SExcel::getRange(const QString &name)
{
    if (mActiveWorkSheet)
        return mActiveWorkSheet->querySubObject("Range(const QString &)", name);
    else
        return NULL;
}

QAxObject *SExcel::getRange(const int startRow, const int startColumn, const int endRow, const int endColumn)
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

void SExcel::setRangeProperty(const QString &rangeName, const QString &propertyName, const QVariant &propertyValue)
{
    QAxObject *range = getRange(rangeName);
    if (range) {
        QByteArray array = propertyName.toLocal8Bit();
        range->setProperty(array.constData(), propertyValue);
        delete range;
    }
}

void SExcel::setRangeProperty(const int startRow, const int startColumn, const int endRow, const int endColumn, const QString &propertyName, const QVariant &propertyValue)
{
    QAxObject *range = getRange(startRow, startColumn, endRow, endColumn);
    if (range) {
        QByteArray array = propertyName.toLocal8Bit();
        range->setProperty(array.constData(), propertyValue);
        delete range;
    }
}

QString SExcel::getRowsName(const int startRow, const int endRow)
{
    return QString("%1:%2").arg(startRow).arg(endRow);
}

QVariant SExcel::getRangeProperty(const QString &rangeName, const QString &propertyName)
{
    QVariant value;
    QAxObject *range = getRange(rangeName);

    if (range) {
        QByteArray array = propertyName.toLocal8Bit();
        value = range->property(array.constData());
        delete range;
    }

    return value;
}

QAxObject *SExcel::getRows(const QString &name)
{
    if (mActiveWorkBook)
        return mActiveWorkSheet->querySubObject("Rows(const QString &)", name);
    else
        return NULL;
}

QAxObject *SExcel::getRows(const int startRow, const int endRow)
{
    return getRows(
                QString("%1:%2").arg(startRow).arg(endRow)
                );
}

QVariant SExcel::getRowsProperty(const QString &rowsName, const QString &propertyName)
{
    QVariant value;
    QAxObject *rows = getRows(rowsName);
    if (rows) {
        QByteArray array = propertyName.toLocal8Bit();
        value = rows->property(array.constData());
        delete rows;
    }
    return value;
}

QVariant SExcel::getRowsProperty(const int startRow, const int endRow, const QString &propertyName)
{
    QVariant value;
    QAxObject *rows = getRows(startRow, endRow);
    if (rows) {
        QByteArray array = propertyName.toLocal8Bit();
        value = rows->property(array.constData());
        delete rows;
    }
    return value;
}

void SExcel::setRowsProperty(const QString &rowsName, const QString &propertyName, const QVariant &propertyValue)
{
    QAxObject *rows = getRows(rowsName);
    if (rows) {
        QByteArray array = propertyName.toLocal8Bit();
        rows->setProperty(array.constData(), propertyValue);
        delete rows;
    }
}

QVariant SExcel::getRangeProperty(const int startRow, const int startColumn, const int endRow, const int endColumn, const QString &propertyName)
{
    QVariant value;
    QAxObject *range = getRange(startRow, startColumn, endRow, endColumn);

    if (range) {
        QByteArray array = propertyName.toLocal8Bit();
        value = range->property(array.constData());
        delete range;
    }

    return value;
}

void SExcel::setRowsProperty(const int startRow, const int endRow, const QString &propertyName, const QVariant &propertyValue)
{
    QAxObject *rows = getRows(startRow, endRow);
    if (rows) {
        QByteArray array = propertyName.toLocal8Bit();
        rows->setProperty(array.constData(), propertyValue);
        delete rows;
    }
}

QString SExcel::getColumnsName(const int startColumn, const int endColumn)
{
    return QString("%1:%2")
            .arg((char)(startColumn - 1 + 'A'))
            .arg((char)(endColumn - 1 + 'A'));
}

QAxObject *SExcel::getColumns(const QString &name)
{
    if (mActiveWorkSheet)
        return mActiveWorkSheet->querySubObject("Columns(const QString &)", name);
    else
        return NULL;
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

QVariant SExcel::getCellProperty(const int row, const int column, const QString &propertyName)
{
    QAxObject *obj = getCell(row, column);
    QByteArray arrayPropertyName = propertyName.toLocal8Bit();
    QVariant v = obj->property(arrayPropertyName.constData());
    delete obj;
    return v;
}

QVariant SExcel::getCellValue(const int row, const int column)
{
    return getCellProperty(row, column, "Value");
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
    if (mActiveWorkSheet) {
        delete mActiveWorkSheet;
        mActiveWorkSheet = NULL;
    }

    if (mActiveWorkBook) {
        delete mActiveWorkBook;
        mActiveWorkBook = NULL;
    }

    if (mWorkBooks) {
        delete mWorkBooks;
        mWorkBooks = NULL;
    }

    if (mExcelApp) {
        delete mExcelApp;
        mExcelApp = NULL;
    }
}



