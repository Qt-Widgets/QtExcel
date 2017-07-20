/**
 * �ļ���: sexcel.h
 *
 * SExcel����һ������Qt��Excel�ĵ�������,
 * �����ṩ�˶�Excel���½�,��,�޸�,����Ȳ���.
 *
 * ��Ҫ��Qt��Ŀ�ļ��м���:
 * QT       += axcontainer
 *
 * VBA Excel�еĶ���:
 * - Ӧ�ö���(Excel.Application): �򿪵�ExcelӦ�ó���ʵ��.
 *
 * - ����������(WorkBooks): ����ExcelӦ���д򿪵����й�����.
 *
 * - ���������(WorkSheets): �����������е����й�����.
 *
 * - ��Χ����(Range): �������е�һ����Χ,������һ����Ԫ��,Ҳ�����Ƕ����Ԫ��.
 *
 * ע��:
 * 1.���к����е�Row��Column���Ǵ�1��ʼ,���Ǵ�0��ʼ.
 */


#ifndef SEXCEL_H
#define SEXCEL_H

#include <QObject>
#include <QAxObject>

class Rows {
public:
    Rows(const int begin, const int end) {
        this->begin = begin;
        this->end = end;
    }
    int begin;
    int end;
};

class Columns {
public:
    Columns(const int begin, const int end) {
        this->begin = begin;
        this->end = end;
    }
    int begin;
    int end;
};

class Range {
public:
    Range(const int beginRow, const int beginColumn, const int endRow, const int endColumn) {
        this->beginRow = beginRow;
        this->beginColumn = beginColumn;
        this->endRow = endRow;
        this->endColumn = endColumn;
    }
    int beginRow;
    int beginColumn;
    int endRow;
    int endColumn;
};

class SExcel : public QObject
{
    Q_OBJECT

public:
    // ˮƽ���뷽ʽ
    enum HAlignment {
        HAlignCenter = -4108, // xlCenter = -4108
        HAlignLeft = -4131,     // xlLeft = -4131
        HAlignRight = -4152    // xlRight = -4152
    };

    // ��ֱ���뷽ʽ
    enum VAlignment {
        VAlignCenter = -4108, // xlCenter = -4108
        VAlignTop = -4106,      // xlTop = -4106
        VAlignBottom = -4107    // xlBottom = -4107
    };

    SExcel(QObject *parent = 0);
    virtual ~SExcel();

    ///
    /// Ӧ�ö���
    ///
    bool execute();                                             /* ����ExcelӦ�ó��� */
    void quit();                                                   /* �ر�ExcelӦ�ó��� */
    inline bool isVisible() const { return mIsVisible; }    /* ��ȡExcelӦ�ó��򴰿��Ƿ�ɼ� */
    inline void setVisible(const bool b) { mIsVisible = b; mExcelApp->setProperty("Visible", mIsVisible); } /* ����ExcelӦ�ó��򴰿��Ƿ�ɼ� */
    inline QString errorString() const { return mErrorString; }     /* ��ȡExcelӦ�ó���Ĵ������� */
    inline bool isExecuted() const { return mIsExecuted; }          /* ��ȡExcelӦ�ó����Ƿ��Ѿ����� */

    ///
    /// ����������
    ///
    int workBooksCount();                                       /* �򿪵Ĺ����������� */
    void newWorkBook();                                         /* �½�������,�½��Ĺ�����������,��Ϊ������� */
    void closeWorkBooks();                                      /* �ر����й����� */
    void saveAsXLS(const QString &filePath);          /* �������������Ϊ��׺��Ϊ xls ��excel�ĵ��ļ� */
    void saveAsXLSX(const QString &filePath);        /* �������������Ϊ��׺��Ϊ xlsx ��excel�ĵ��ļ� */
    void open(const QString &filePath);                    /* ���µĹ������д��ļ� */
    void save();                                                         /* �������������޸� */

    ///
    /// ���������
    ///
    void setActiveWorkSheet(const int index);                                     /* ���û������ */
    int workSheetsCount();                                                                  /* ��ȡ��������еĹ��������� */
    void setActiveWorkSheetName(const QString &name);                   /* ���û����������� */
    void setWorkSheetName(const int index, const QString &name);    /* ���û�������е�index������������� */

    ///
    /// ��Χ����,����һ����Χ�ڵ����е�Ԫ��
    ///
    static QString getRangeName(const int startRow, const int startColumn, const int endRow, const int endColumn);
    QAxObject *getRange(const QString &name);
    QAxObject *getRange(const int startRow, const int startColumn, const int endRow, const int endColumn);
    QVariant getRangeProperty(const QString &rangeName, const QString &propertyName);
    QVariant getRangeProperty(const int startRow, const int startColumn, const int endRow, const int endColumn, const QString &propertyName);
    void setRangeProperty(const QString &rangeName, const QString &propertyName, const QVariant &propertyValue);
    void setRangeProperty(const int startRow, const int startColumn, const int endRow, const int endColumn, const QString &propertyName, const QVariant &propertyValue);

    ///
    /// �ж���,����ĳһ�л��߶��е����е�Ԫ��
    ///
    static QString getRowsName(const int startRow, const int endRow);
    QAxObject *getRows(const QString &name);
    QAxObject *getRows(const int startRow, const int endRow);
    QVariant getRowsProperty(const QString &rowsName, const QString &propertyName);
    QVariant getRowsProperty(const int startRow, const int endRow, const QString &propertyName);
    void setRowsProperty(const QString &rowsName, const QString &propertyName, const QVariant &propertyValue);
    void setRowsProperty(const int startRow, const int endRow, const QString &propertyName, const QVariant &propertyValue);

    ///
    /// �ж���,����ĳһ�л��߶��е����е�Ԫ��
    ///
    static QString getColumnsName(const int startColumn, const int endColumn);
    QAxObject *getColumns(const QString &name);
    QAxObject *getColumns(const int startColumn, const int endColumn);
    void setColumnsProperty(const int startColumn, const int endColumn, const QString &propertyName, const QVariant &propertyValue);

    ///
    /// ��Ԫ�����
    ///
    QAxObject *getCell(const int row, const int column);
    void setCellProperty(const int row, const int column, const QString &propertyName, const QVariant &propertyValue);
    void setCellText(const int row, const int column, const QString &text);

    QVariant getCellProperty(const int row, const int column, const QString &propertyName);
    QVariant getCellValue(const int row, const int column);

protected:
    bool initCOM();                             /* ��ʼ��windows COM��� */
    void deleteAxObjects();

private:
    QAxObject *mExcelApp;               /* excelӦ�ó���ʵ�� */
    QAxObject *mWorkBooks;              /* ���еĹ�����,excelӦ��һ����,��������й����� */
    QAxObject *mActiveWorkBook;     /* �������,������ļ������½���������,�Ż��û������ */
    QAxObject *mActiveWorkSheet;   /* �������,�����û��������,���ܻ�û�Ĺ����� */

    bool            mIsValid;                   /* ���ĳЩԭ���¶����ʼ��ʧ��,���������ϵ�Excel������˵�,��ô�����������Ч�� */
    QString     mErrorString;               /* excel����Ĵ������� */

    bool            mIsVisible;                 /* �򿪵�ExcelӦ�ó����Ƿ�ɼ� */
    bool            mIsExecuted;            /* excelӦ���Ƿ��Ѿ����� */
};

#endif // SEXCEL_H
