#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QAxObject>
#include <QAxWidget>
#include <QFileDialog>



MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);     
         QString filepath=QFileDialog::getSaveFileName(NULL,QObject::tr("Save orbit"),"/untitled.xls",QObject::tr("/*Microsoft Office 2016 */(*.xlsx)"));//获取保存路径
         QList<QVariant> allRowsData;//保存所有行数据
         allRowsData.clear();
     //    mLstData.append(QVariant(12));
         if(!filepath.isEmpty()){
             QAxObject *excel = new QAxObject("Excel.Application");//连接Excel控件
             excel->dynamicCall("SetVisible (bool Visible)",false);//不显示窗体
             excel->setProperty("DisplayAlerts", true);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
             QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
             workbooks->dynamicCall("Add");//新建一个工作簿
             QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
             QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
             QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1

             for(int row = 1; row <= 1000; row++)
             {
                 QList<QVariant> aRowData;//保存一行数据
                 for(int column = 1; column <= 2; column++)
                 {
                     aRowData.append(QVariant(/*row*column*/"test"));
                 }
                 allRowsData.append(QVariant(aRowData));
             }

             QAxObject *range = worksheet->querySubObject("Range(const QString )", "A1:B1000");
             range->dynamicCall("SetValue(const QVariant&)",QVariant(allRowsData));//存储所有数据到 excel 中,批量操作,速度极快
             range->querySubObject("Font")->setProperty("Size", 30);//设置字号

             QAxObject *cell = worksheet->querySubObject("Range(QVariant, QVariant)","A1");//获取单元格
             cell = worksheet->querySubObject("Cells(int, int)", 1, 1);//等同于上一句
             cell->dynamicCall("SetValue(const QVariant&)",QVariant(123));//存储一个 int 数据到 excel 的单元格中
             cell->dynamicCall("SetValue(const QVariant&)",QVariant("abc"));//存储一个 string 数据到 excel 的单元格中

             QString str = cell->dynamicCall("Value2()").toString();//读取单元格中的值

             /*QAxObject *font = cell->querySubObject("Font");
             font->setProperty("Name", itemFont.family());  //设置单元格字体
             font->setProperty("Bold", itemFont.bold());  //设置单元格字体加粗
             font->setProperty("Size", itemFont.pixelSize());  //设置单元格字体大小
             font->setProperty("Italic",itemFont.italic());  //设置单元格字体斜体
             font->setProperty("Underline", itemFont.underline());  //设置单元格下划线
             font->setProperty("Color", item->foreground().color());  //设置单元格字体颜色*/
             worksheet->querySubObject("Range(const QString&)", "1:1")->setProperty("RowHeight", 60);//调整第一行行高

             workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));//保存至filepath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
             workbook->dynamicCall("Close()");//关闭工作簿
             excel->dynamicCall("Quit()");//关闭excel
             delete excel;
             excel=NULL;
         }
}

MainWindow::~MainWindow()
{
    delete ui;
}
