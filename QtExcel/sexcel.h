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

    // ���ݸ�ʽ


    SExcel(QObject *parent = 0);
    virtual ~SExcel();

    inline QString errorString() const { return mErrorString; }
    inline bool isVisible() const { return mIsVisible; }
    inline void setVisible(const bool b) { mIsVisible = b; mExcelApp->setProperty("Visible", mIsVisible); }
    inline int workBooksCount() const { return mWorkBooksCount; }
    inline bool isAppExecuted() const { return mIsAppExecuted; }

    // Ӧ�ö���
    bool executeApp();                                             /* ����ExcelӦ�ó��� */
    void quitApp();                                                   /* �ر�ExcelӦ�ó��� */

    // ����������
    void newWorkBook();                                         /* �½������� */
    void closeWorkBooks();                                      /* �ر����й����� */
    void saveAsXLS(const QString &filePath);          /* �������������Ϊ��׺��Ϊ xls ��excel�ĵ��ļ� */
    void saveAsXLSX(const QString &filePath);        /* �������������Ϊ��׺��Ϊ xlsx ��excel�ĵ��ļ� */
    void open(const QString &filePath);                    /* ���µĹ������д��ļ� */
    void save();                                                         /* �������������޸� */

    // ���������
    void workSheetsCount();                                                                 /* ��������еĹ��������� */
    bool setActiveWorkSheetName(const QString &name);                   /* ���õ�ǰ����������� */
    bool setWorkSheetName(const int index, const QString &name);    /* ���û�������е�index������������� */
    bool setActiveWorkSheet(const int index);                                     /* ���û������ */

    // ��Χ����,����һ����Χ�ڵ����е�Ԫ��
    QAxObject *getRange(const QString &name);                                                                                                                             /* ��ȡ��Χ���� */
    QAxObject *getRange(const int startRow, const int startColumn, const int endRow, const int endColumn);                                    /* ��ȡ��Χ���� */
    void setRangeHAlignment(const int startRow, const int startColumn, const int endRow, const int endColumn, HAlignment align);    /* ���÷�Χ�����ˮƽ���뷽ʽ */
    void setRangeVAlignment(const int startRow, const int startColumn, const int endRow, const int endColumn, VAlignment align);    /* ���÷�Χ����Ĵ�ֱ���뷽ʽ */
    void setRangeMergeCells(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b);            /* ���÷�Χ�����Ƿ�ϲ���Ԫ�� */
    void setRangeWrapText(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b);              /* ���÷�Χ�����Ƿ��Զ����� */
    //  �������,
    // .Name = "����"
    // .FontStyle = "�Ӵ���б"
    // .Size = 12
    // .Strikethrough = True
    // .Superscript = False
    // .Subscript = False
    // .OutlineFont = False
    // .Shadow = False
    // .Underline = xlUnderlineStyleNone
    // .ThemeColor = xlThemeColorLight1
    // .TintAndShade = 0
    // .ThemeFont = xlThemeFontMajor
    //void setRangeNumberFormatLocal(const int startRow, const int startColumn, const int endRow, const int endColumn, const QString &numberFormatLocal);

    // �ж���,����ĳһ�л��߶��е����е�Ԫ��
    QAxObject *getRows(const QString &name);                                                         /* ��ȡ�ж��� */
    QAxObject *getRows(const int startRow, const int endRow);                                   /* ��ȡ�ж��� */
    void setRowsHeight(const int startRow, const int endRow, const float height);        /* �����и� */
    void setRowsHAlignment(const int startRow, const int endRow, HAlignment align); /* ������ˮƽ���뷽ʽ */
    void setRowsVAlignment(const int startRow, const int endRow, VAlignment align); /* �����д�ֱ���뷽ʽ */

    // �ж���,����ĳһ�л��߶��е����е�Ԫ��
    QAxObject *getColumns(const QString &name);                                                             /* ��ȡ�ж��� */
    QAxObject *getColumns(const int startColumn, const int endColumn);                              /* ��ȡ�ж��� */
    void setColumnsWidth(const int startColumn, const int endColumn, const float width);    /* �����п� */
    void setColumnsHAlignment(const int startColumn, const int endColumn, HAlignment align); /* ������ˮƽ���뷽ʽ */
    void setColumnsVAlignment(const int startColumn, const int endColumn, VAlignment align); /* �����д�ֱ���뷽ʽ */

    // ��Ԫ�����
    QAxObject *getCell(const int row, const int column);
    void setCellText(const int row, const int column, const QString &text);

protected:
    bool initCOM();                             /* ��ʼ��windows COM��� */
    void deleteAxObjects();

private:
    QAxObject *mExcelApp;               /* excelӦ�ó���ʵ�� */
    QAxObject *mWorkBooks;              /* ���еĹ����� */
    int                 mWorkBooksCount;    /* ���������� */
    QAxObject *mActiveWorkBook;     /* ������� */
    QAxObject *mActiveWorkSheet;   /* ������� */

    bool            mIsValid;                   /* ���ĳЩԭ���¶����ʼ��ʧ��,���������ϵ�Excel������˵�,��ô�����������Ч�� */
    QString     mErrorString;               /* excel����Ĵ������� */

    bool            mIsVisible;                 /* �򿪵�ExcelӦ�ó����Ƿ�ɼ� */
    bool            mIsAppExecuted;            /* excelӦ���Ƿ��Ѿ����� */
};

#endif // SEXCEL_H
