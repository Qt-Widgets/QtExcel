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
    if (mActiveWorkBook) {
        mActiveWorkBook->dynamicCall("SaveAs(const QString &, int, const QString &, const QString &, bool, bool)",
                                    QDir::toNativeSeparators(filePath), 56, "", "", false, false);
    }
}

void SExcel::saveAsXLSX(const QString &filePath)
{
    if (mActiveWorkBook) {
        mActiveWorkBook->dynamicCall("SaveAs(const QString &, int, const QString &, const QString &, bool, bool)",
                                    QDir::toNativeSeparators(filePath), 51, "", "", false, false);
    }
}

void SExcel::open(const QString &filePath)
{
    if (!mIsExecuted)
        return;

    mWorkBooks->dynamicCall("Open(const QString &)", filePath);
    mActiveWorkBook = mExcelApp->querySubObject("ActiveWorkBook");
    mActiveWorkSheet = mActiveWorkBook->querySubObject("ActiveSheet");
}

void SExcel::save()
{
    if (!mIsExecuted)
        return;

    mActiveWorkBook->dynamicCall("Save()");
}

void SExcel::setActiveWorkSheetProperty(const QString &propertyName, const QVariant &value)
{
    if (mActiveWorkSheet) {
        QByteArray array = propertyName.toLocal8Bit();
        mActiveWorkSheet->setProperty(array.constData(), value);
    }
}

void SExcel::setActiveWorkSheet(const int index)
{
    int count = workSheetsCount();
    if (index <= 0 || count <= 0 || index > count)
        return;

    if (mActiveWorkBook) {
        QAxObject *workSheet = mActiveWorkBook->querySubObject("WorkSheets(int)", index);
        workSheet->dynamicCall("Activate(void)");
        mActiveWorkSheet = workSheet;
    }

}

int SExcel::workSheetsCount()
{
    if (mActiveWorkBook) {
        QAxObject *workSheets = mActiveWorkBook->querySubObject("WorkSheets");
        int count = workSheets->property("Count").toInt();
        delete workSheets;
        return count;
    } else {
        return 0;
    }
}

void SExcel::newWorkSheet()
{
    if (mActiveWorkBook) {
        QAxObject *workSheets = mActiveWorkBook->querySubObject("WorkSheets");
        workSheets->dynamicCall("Add(void)");
        delete workSheets;
        mActiveWorkSheet = mActiveWorkBook->querySubObject("ActiveSheet");
    }
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

QVariant SExcel::getRangeProperty(const int startRow, const int startColumn, const int endRow, const int endColumn, const QString &propertyName)
{
    return getRangeProperty(getRangeName(startRow, startColumn, endRow, endColumn), propertyName);
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
    setRangeProperty(getRangeName(startRow, startColumn, endRow, endColumn), propertyName, propertyValue);
}

bool SExcel::execute()
{
    // ����Ѿ�����
    if (mIsExecuted)
        return true;

    // ��excel����
    mExcelApp = new QAxObject("Excel.Application");
    if (mExcelApp->isNull()) {
        mErrorString = QString::fromLocal8Bit("�޷���ExcelӦ�ó���,��ȷ�����ĵ����Ѿ���ȷ��װ��Excel");
        mIsValid = false;
        return false;
    }

    // ����app���ڿɼ���
    mExcelApp->setProperty("Visible", mIsVisible);

    // �õ����й���������
    mWorkBooks = mExcelApp->querySubObject("WorkBooks");

    // ���ñ�־λ
    mIsExecuted = true;

    return true;
}

void SExcel::quit()
{
    // �ر�Ӧ��,�ͷ��ڴ�,����QAxObject������ΪNULL
    if (mIsExecuted) {
        mExcelApp->dynamicCall("Quit(void)");
        deleteAxObjects();
        mIsExecuted = false;
    }
}

void SExcel::setVisible(const bool b)
{
    mIsVisible = b;
    if (mExcelApp)
        mExcelApp->setProperty("Visible", mIsVisible);
}

void SExcel::setActiveWorkBookProperty(const QString &propertyName, const QVariant &propertyValue)
{
    if (mActiveWorkBook) {
        QByteArray array = propertyName.toLocal8Bit();
        mActiveWorkBook->setProperty(array.constData(), propertyValue);
    }
}

void SExcel::setActiveWorkBook(const int index)
{
    if (index <= 0)
        return;

    int count = workBooksCount();
    if (index <= count) {
        if (mExcelApp) {
            QAxObject *workBook = mExcelApp->querySubObject("WorkBooks(int)", index);
            workBook->dynamicCall("Activate(void)");
            mActiveWorkBook = workBook;
            mActiveWorkSheet = mActiveWorkBook->querySubObject("ActiveSheet");
        }
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

    // �����Ĺ�����������,��Ϊ�������
    mActiveWorkBook = mExcelApp->querySubObject("ActiveWorkBook");

    // ��������ĵ�һ��������Ĭ��Ϊ�������
    mActiveWorkSheet = mActiveWorkBook->querySubObject("ActiveSheet");
}

void SExcel::closeWorkBooks()
{
    // �ر����д򿪵Ĺ�����,
    // �ͷŻ�������ͻ������,��ΪNULL
    if (mWorkBooks) {
        mWorkBooks->dynamicCall("Close");
        delete mActiveWorkSheet;
        mActiveWorkSheet = NULL;
        delete mActiveWorkBook;
        mActiveWorkBook = NULL;
    }
}

QString SExcel::getRowsName(const int startRow, const int endRow)
{
    return QString("%1:%2").arg(startRow).arg(endRow);
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
    return getRowsProperty(getRowsName(startRow, endRow), propertyName);
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

void SExcel::setRowsProperty(const int startRow, const int endRow, const QString &propertyName, const QVariant &propertyValue)
{
    setRowsProperty(getRowsName(startRow, endRow), propertyName, propertyValue);
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

QVariant SExcel::getColumnsProperty(const QString &columnsName, const QString &propertyName)
{
    QVariant value;
    QAxObject *columns = getColumns(columnsName);
    if (columns) {
        QByteArray array = propertyName.toLocal8Bit();
        value = columns->property(array.constData());
        delete columns;
    }
    return value;
}

QVariant SExcel::getColumnsProperty(const int startColumn, const int endColumn, const QString &propertyName)
{
    return getColumnsProperty(getColumnsName(startColumn, endColumn), propertyName);
}

void SExcel::setColumnsProperty(const QString &columnsName, const QString &propertyName, const QVariant &propertyValue)
{
    QAxObject *columns = getColumns(columnsName);
    if (columns) {
        QByteArray array = propertyName.toLocal8Bit();
        columns->setProperty(array.constData(), propertyValue);
        delete columns;
    }
}

void SExcel::setColumnsProperty(const int startColumn, const int endColumn, const QString &propertyName, const QVariant &propertyValue)
{
    setColumnsProperty(getColumnsName(startColumn, endColumn), propertyName, propertyValue);
}

QAxObject *SExcel::getCell(const int row, const int column)
{
    return mActiveWorkSheet->querySubObject("Cells(int,int)", row, column);
}

void SExcel::setCellProperty(const int row, const int column, const QString &propertyName, const QVariant &propertyValue)
{
    QAxObject *cell = getCell(row, column);
    QByteArray array = propertyName.toLocal8Bit();
    cell->setProperty(array.constData(), propertyValue);
    delete cell;
}

QVariant SExcel::getCellProperty(const int row, const int column, const QString &propertyName)
{
    QVariant value;
    QAxObject *cell = getCell(row, column);
    if (cell) {
        QByteArray array = propertyName.toLocal8Bit();
        value = cell->property(array.constData());
        delete cell;
    }
    return value;
}

bool SExcel::initCOM()
{
    HRESULT r = OleInitialize(0);
    if (S_OK != r && S_FALSE != r) {
        mErrorString = QString::fromLocal8Bit("�޷���ʼ��Windows COM���");
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



