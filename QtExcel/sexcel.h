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

/*
HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        */

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
    bool isVisible() const { return mIsVisible; }    /* ��ȡExcelӦ�ó��򴰿��Ƿ�ɼ� */
    void setVisible(const bool b);                           /* ����ExcelӦ�ó��򴰿��Ƿ�ɼ� */
    QString errorString() const { return mErrorString; }     /* ��ȡExcelӦ�ó���Ĵ������� */
    bool isExecuted() const { return mIsExecuted; }          /* ��ȡExcelӦ�ó����Ƿ��Ѿ����� */

    ///
    /// ����������
    ///
    void setActiveWorkBookProperty(const QString &propertyName, const QVariant &propertyValue);
    void setActiveWorkBook(const int index);          /* ���û������ */
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
    void setActiveWorkSheetProperty(const QString &propertyName, const QVariant &value);
    void setActiveWorkSheet(const int index);                                     /* ���û������ */
    int workSheetsCount();                                                                  /* ��ȡ��������еĹ��������� */
    void newWorkSheet();                                                                    /* �½�һ�������� */

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
    QVariant getColumnsProperty(const QString &columnsName, const QString &propertyName);
    QVariant getColumnsProperty(const int startColumn, const int endColumn, const QString &propertyName);
    void setColumnsProperty(const QString &columnsName, const QString &propertyName, const QVariant &propertyValue);
    void setColumnsProperty(const int startColumn, const int endColumn, const QString &propertyName, const QVariant &propertyValue);

    ///
    /// ��Ԫ�����
    ///
    QAxObject *getCell(const int row, const int column);
    void setCellProperty(const int row, const int column, const QString &propertyName, const QVariant &propertyValue);
    QVariant getCellProperty(const int row, const int column, const QString &propertyName);

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
