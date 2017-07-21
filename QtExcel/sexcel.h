/**
 * 文件名: sexcel.h
 *
 * SExcel类是一个基于Qt的Excel文档操作类,
 * 该类提供了对Excel的新建,打开,修改,保存等操作.
 *
 * 需要在Qt项目文件中加入:
 * QT       += axcontainer
 *
 * VBA Excel中的对象:
 * - 应用对象(Excel.Application): 打开的Excel应用程序实例.
 *
 * - 工作簿对象(WorkBooks): 包含Excel应用中打开的所有工作簿.
 *
 * - 工作表对象(WorkSheets): 包含工作簿中的所有工作表.
 *
 * - 范围对象(Range): 工作表中的一定范围,可以是一个单元格,也可以是多个单元格.
 *
 * 注意:
 * 1.所有函数中的Row和Column都是从1开始,不是从0开始.
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
    // 水平对齐方式
    enum HAlignment {
        HAlignCenter = -4108, // xlCenter = -4108
        HAlignLeft = -4131,     // xlLeft = -4131
        HAlignRight = -4152    // xlRight = -4152
    };

    // 垂直对齐方式
    enum VAlignment {
        VAlignCenter = -4108, // xlCenter = -4108
        VAlignTop = -4106,      // xlTop = -4106
        VAlignBottom = -4107    // xlBottom = -4107
    };

    SExcel(QObject *parent = 0);
    virtual ~SExcel();

    ///
    /// 应用对象
    ///
    bool execute();                                             /* 运行Excel应用程序 */
    void quit();                                                   /* 关闭Excel应用程序 */
    bool isVisible() const { return mIsVisible; }    /* 获取Excel应用程序窗口是否可见 */
    void setVisible(const bool b);                           /* 设置Excel应用程序窗口是否可见 */
    QString errorString() const { return mErrorString; }     /* 获取Excel应用程序的错误描述 */
    bool isExecuted() const { return mIsExecuted; }          /* 获取Excel应用程序是否已经启动 */

    ///
    /// 工作簿对象
    ///
    void setActiveWorkBookProperty(const QString &propertyName, const QVariant &propertyValue);
    void setActiveWorkBook(const int index);          /* 设置活动工作簿 */
    int workBooksCount();                                       /* 打开的工作簿的总数 */
    void newWorkBook();                                         /* 新建工作簿,新建的工作簿被激活,成为活动工作簿 */
    void closeWorkBooks();                                      /* 关闭所有工作簿 */
    void saveAsXLS(const QString &filePath);          /* 将活动工作簿保存为后缀名为 xls 的excel文档文件 */
    void saveAsXLSX(const QString &filePath);        /* 将活动工作簿保存为后缀名为 xlsx 的excel文档文件 */
    void open(const QString &filePath);                    /* 在新的工作簿中打开文件 */
    void save();                                                         /* 保存活动工作簿的修改 */

    ///
    /// 工作表对象
    ///
    void setActiveWorkSheetProperty(const QString &propertyName, const QVariant &value);
    void setActiveWorkSheet(const int index);                                     /* 设置活动工作表 */
    int workSheetsCount();                                                                  /* 获取活动工作簿中的工作表总数 */
    void newWorkSheet();                                                                    /* 新建一个工作表 */

    ///
    /// 范围对象,包含一定范围内的所有单元格
    ///
    static QString getRangeName(const int startRow, const int startColumn, const int endRow, const int endColumn);
    QAxObject *getRange(const QString &name);
    QAxObject *getRange(const int startRow, const int startColumn, const int endRow, const int endColumn);
    QVariant getRangeProperty(const QString &rangeName, const QString &propertyName);
    QVariant getRangeProperty(const int startRow, const int startColumn, const int endRow, const int endColumn, const QString &propertyName);
    void setRangeProperty(const QString &rangeName, const QString &propertyName, const QVariant &propertyValue);
    void setRangeProperty(const int startRow, const int startColumn, const int endRow, const int endColumn, const QString &propertyName, const QVariant &propertyValue);

    ///
    /// 行对象,包含某一行或者多行的所有单元格
    ///
    static QString getRowsName(const int startRow, const int endRow);
    QAxObject *getRows(const QString &name);
    QAxObject *getRows(const int startRow, const int endRow);
    QVariant getRowsProperty(const QString &rowsName, const QString &propertyName);
    QVariant getRowsProperty(const int startRow, const int endRow, const QString &propertyName);
    void setRowsProperty(const QString &rowsName, const QString &propertyName, const QVariant &propertyValue);
    void setRowsProperty(const int startRow, const int endRow, const QString &propertyName, const QVariant &propertyValue);

    ///
    /// 列对象,包含某一列或者多列的所有单元格
    ///
    static QString getColumnsName(const int startColumn, const int endColumn);
    QAxObject *getColumns(const QString &name);
    QAxObject *getColumns(const int startColumn, const int endColumn);
    QVariant getColumnsProperty(const QString &columnsName, const QString &propertyName);
    QVariant getColumnsProperty(const int startColumn, const int endColumn, const QString &propertyName);
    void setColumnsProperty(const QString &columnsName, const QString &propertyName, const QVariant &propertyValue);
    void setColumnsProperty(const int startColumn, const int endColumn, const QString &propertyName, const QVariant &propertyValue);

    ///
    /// 单元格对象
    ///
    QAxObject *getCell(const int row, const int column);
    void setCellProperty(const int row, const int column, const QString &propertyName, const QVariant &propertyValue);
    QVariant getCellProperty(const int row, const int column, const QString &propertyName);

protected:
    bool initCOM();                             /* 初始化windows COM组件 */
    void deleteAxObjects();

private:
    QAxObject *mExcelApp;               /* excel应用程序实例 */
    QAxObject *mWorkBooks;              /* 所有的工作簿,excel应用一启动,便会获得所有工作簿 */
    QAxObject *mActiveWorkBook;     /* 活动工作簿,必须打开文件或者新建工作簿后,才会获得活动工作簿 */
    QAxObject *mActiveWorkSheet;   /* 活动工作表,必须获得活动工作簿后,才能获得活动的工作表 */

    bool            mIsValid;                   /* 如果某些原因导致对象初始化失败,比如计算机上的Excel软件损坏了等,那么这个对象是无效的 */
    QString     mErrorString;               /* excel对象的错误描述 */

    bool            mIsVisible;                 /* 打开的Excel应用程序是否可见 */
    bool            mIsExecuted;            /* excel应用是否已经运行 */
};

#endif // SEXCEL_H
