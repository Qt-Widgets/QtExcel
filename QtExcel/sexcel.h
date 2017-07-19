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

    // 数据格式


    SExcel(QObject *parent = 0);
    virtual ~SExcel();

    inline QString errorString() const { return mErrorString; }
    inline bool isVisible() const { return mIsVisible; }
    inline void setVisible(const bool b) { mIsVisible = b; mExcelApp->setProperty("Visible", mIsVisible); }
    inline int workBooksCount() const { return mWorkBooksCount; }
    inline bool isAppExecuted() const { return mIsAppExecuted; }

    // 应用对象
    bool executeApp();                                             /* 运行Excel应用程序 */
    void quitApp();                                                   /* 关闭Excel应用程序 */

    // 工作簿对象
    void newWorkBook();                                         /* 新建工作簿 */
    void closeWorkBooks();                                      /* 关闭所有工作簿 */
    void saveAsXLS(const QString &filePath);          /* 将活动工作簿保存为后缀名为 xls 的excel文档文件 */
    void saveAsXLSX(const QString &filePath);        /* 将活动工作簿保存为后缀名为 xlsx 的excel文档文件 */
    void open(const QString &filePath);                    /* 在新的工作簿中打开文件 */
    void save();                                                         /* 保存活动工作簿的修改 */

    // 工作表对象
    void workSheetsCount();                                                                 /* 活动工作簿中的工作表总数 */
    bool setActiveWorkSheetName(const QString &name);                   /* 设置当前工作表的名称 */
    bool setWorkSheetName(const int index, const QString &name);    /* 设置活动工作簿中第index个工作表的名称 */
    bool setActiveWorkSheet(const int index);                                     /* 设置活动工作表 */

    // 范围对象,包含一定范围内的所有单元格
    QAxObject *getRange(const QString &name);                                                                                                                             /* 获取范围对象 */
    QAxObject *getRange(const int startRow, const int startColumn, const int endRow, const int endColumn);                                    /* 获取范围对象 */
    void setRangeHAlignment(const int startRow, const int startColumn, const int endRow, const int endColumn, HAlignment align);    /* 设置范围对象的水平对齐方式 */
    void setRangeVAlignment(const int startRow, const int startColumn, const int endRow, const int endColumn, VAlignment align);    /* 设置范围对象的垂直对齐方式 */
    void setRangeMergeCells(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b);            /* 设置范围对象是否合并单元格 */
    void setRangeWrapText(const int startRow, const int startColumn, const int endRow, const int endColumn, const bool b);              /* 设置范围对象是否自动换行 */
    //  字体对象,
    // .Name = "宋体"
    // .FontStyle = "加粗倾斜"
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

    // 行对象,包含某一行或者多行的所有单元格
    QAxObject *getRows(const QString &name);                                                         /* 获取行对象 */
    QAxObject *getRows(const int startRow, const int endRow);                                   /* 获取行对象 */
    void setRowsHeight(const int startRow, const int endRow, const float height);        /* 设置行高 */
    void setRowsHAlignment(const int startRow, const int endRow, HAlignment align); /* 设置行水平对齐方式 */
    void setRowsVAlignment(const int startRow, const int endRow, VAlignment align); /* 设置行垂直对齐方式 */

    // 列对象,包含某一列或者多列的所有单元格
    QAxObject *getColumns(const QString &name);                                                             /* 获取列对象 */
    QAxObject *getColumns(const int startColumn, const int endColumn);                              /* 获取列对象 */
    void setColumnsWidth(const int startColumn, const int endColumn, const float width);    /* 设置列宽 */
    void setColumnsHAlignment(const int startColumn, const int endColumn, HAlignment align); /* 设置列水平对齐方式 */
    void setColumnsVAlignment(const int startColumn, const int endColumn, VAlignment align); /* 设置列垂直对齐方式 */

    // 单元格对象
    QAxObject *getCell(const int row, const int column);
    void setCellText(const int row, const int column, const QString &text);

protected:
    bool initCOM();                             /* 初始化windows COM组件 */
    void deleteAxObjects();

private:
    QAxObject *mExcelApp;               /* excel应用程序实例 */
    QAxObject *mWorkBooks;              /* 所有的工作簿 */
    int                 mWorkBooksCount;    /* 工作簿总数 */
    QAxObject *mActiveWorkBook;     /* 活动工作簿 */
    QAxObject *mActiveWorkSheet;   /* 活动工作表 */

    bool            mIsValid;                   /* 如果某些原因导致对象初始化失败,比如计算机上的Excel软件损坏了等,那么这个对象是无效的 */
    QString     mErrorString;               /* excel对象的错误描述 */

    bool            mIsVisible;                 /* 打开的Excel应用程序是否可见 */
    bool            mIsAppExecuted;            /* excel应用是否已经运行 */
};

#endif // SEXCEL_H
