#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QFileDialog>
#include <QDebug>
#include <QStringList>

#include <minwindef.h>
#include "psapi.h"

#include <processthreadsapi.h>



MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ui->tableWidget_name->setColumnCount(2);
    ui->lineEdit_workSheetName->setText(QString("CMS"));
    ui->lineEdit_defaultLanguage->setText(QString("English"));
    ui->lineEdit_defaultKeyName->setText(QString("Key"));
    ui->lineEdit_defaultConfigPath->setText(QString("%1").arg(QCoreApplication::applicationDirPath()+"/"+"config_2.ini"));
    set_connect();
}

MainWindow::~MainWindow()
{

    bool flag = FindProcess(QString("C:\\PROGRAM FILES (X86)\\MICROSOFT OFFICE\\ROOT\\OFFICE16\\EXCEL.EXE").toStdString().c_str());
    while(flag)
    {
        CloseProcess("C:\\PROGRAM FILES (X86)\\MICROSOFT OFFICE\\ROOT\\OFFICE16\\EXCEL.EXE");
        flag =FindProcess("C:\\PROGRAM FILES (X86)\\MICROSOFT OFFICE\\ROOT\\OFFICE16\\EXCEL.EXE");
    }


    delete ui;
}

void MainWindow::on_open_file(bool)
{
    QString file_path = QFileDialog::getOpenFileName(nullptr,QString("open"));
    ui->lineEdit_filePath->setText(file_path);
    //    get_tar_work_sheet(workbook,)
}

void MainWindow::load_ini()
{
    QString ini_path = ui->lineEdit_defaultConfigPath->text();
    QFile file(ini_path);
    if(file.open(QFile::ReadOnly|QIODevice::Unbuffered |QIODevice::Text))
    {
        QTextStream in(&file);
        in.setCodec("UTF-8");
        QString line;
        ui->tableWidget_name->clear();
        ui->tableWidget_name->setRowCount(0);
        while(!in.atEnd())
        {
            line = in.readLine();
            line.remove("\n");
            QStringList list = line.split("=>");
            if(list.count()<2)
                continue;
            QString str_0 = list.at(0);
            QString str_0_noEmpty = str_0.trimmed();
            int count = ui->tableWidget_name->rowCount();
            ui->tableWidget_name->insertRow(count);
            ui->tableWidget_name->setItem(count,0,new QTableWidgetItem(str_0_noEmpty));
            ui->tableWidget_name->setItem(count,1,new QTableWidgetItem(list.at(1)));
            ++count;
        }
    }
}

void MainWindow::create_lang()
{
    QAxObject* workbook=read_excel(ui->lineEdit_filePath->text());
    if(!workbook)
        return;
    QAxObject* worksheet = get_tar_work_sheet(workbook,ui->lineEdit_workSheetName->text());
    if(!worksheet)
        return;
    //获取英语所在行数
    int default_language_column = get_tar_sheet_column(worksheet,ui->lineEdit_defaultLanguage->text());
    for(int i=0;i<ui->tableWidget_name->rowCount();++i)
    {

        int tar_column = get_tar_sheet_column(worksheet,ui->tableWidget_name->item(i,0)->text());
        qDebug()<<QString("语言： %1  , 列数： %2").arg(ui->tableWidget_name->item(i,0)->text()).arg(tar_column);
        if(tar_column==-1)
            continue;
        QString file_name = ui->tableWidget_name->item(i,1)->text();
        QString dir_path = QCoreApplication::applicationDirPath();
        QDir tempDir;
        QString file_full_path = dir_path+"/"+file_name;
        if(!tempDir.exists(dir_path))
        {
            tempDir.mkpath(file_full_path);
        }
        QFile tempFile;
        tempDir.setCurrent(file_full_path);
        if(tempFile.exists(file_name))
        {
            qDebug()<<QString("file %1 has already exist").arg(file_name);
            //            continue;
        }
        tempFile.setFileName(file_name);
        if(!tempFile.open(QIODevice::ReadWrite|QIODevice::Text))
            return;
        qint64 pos = tempFile.size();
        tempFile.seek(pos);
        QAxObject *range ;
        QAxObject * usedrange = worksheet->querySubObject("UsedRange");//获取该sheet的使用范围对象

        QAxObject* rows = usedrange->querySubObject("Rows");
        int nRows=rows->property("Count").toInt();
        QString content_prefix;
        content_prefix = QString("[Info]\nLanguage=%1\n[String]\n").arg( ui->tableWidget_name->item(i,0)->text());
        tempFile.write(content_prefix.toUtf8());
        for(int i=2;i<=nRows;++i)
        {

            QString content;
            QString tempStr;
            tempStr.clear();
            content.clear();
            range = worksheet->querySubObject("Cells(int,int)",i,1);
            tempStr= range->dynamicCall("Value2()").toString();
            if(tempStr.isEmpty())
                continue;
            content +=tempStr;
            content +="=";
            tempStr.clear();
            range = worksheet->querySubObject("Cells(int,int)",i,tar_column);
            tempStr= range->dynamicCall("Value2()").toString();
            if(tempStr.isEmpty())
            {

                range = worksheet->querySubObject("Cells(int,int)",i,default_language_column);
                tempStr = range->dynamicCall("Value2()").toString();
                if(tempStr.isEmpty())
                    continue;
            }
            content +=tempStr;
            content+="\n";
            if(content.isEmpty())
                continue;
            //            qDebug()<<QString(content);
            tempFile.write(content.toUtf8());
        }
        tempFile.close();
    }
    workbook->dynamicCall("Close()");
}

void MainWindow::create_lang_morefast()
{
    QAxObject* workbook=read_excel(ui->lineEdit_filePath->text());
    if(!workbook)
        return;
    QAxObject* worksheet = get_tar_work_sheet(workbook,ui->lineEdit_workSheetName->text());
    if(!worksheet)
        return;
    //获取英语所在行数
    int default_language_column = get_tar_sheet_column(worksheet,ui->lineEdit_defaultLanguage->text());
    int default_key_column = get_tar_sheet_column(worksheet,ui->lineEdit_defaultKeyName->text());
    for(int i=0;i<ui->tableWidget_name->rowCount();++i)
    {

        int tar_column = get_tar_sheet_column(worksheet,ui->tableWidget_name->item(i,0)->text());
        qDebug()<<QString("语言： %1  , 列数： %2").arg(ui->tableWidget_name->item(i,0)->text()).arg(tar_column);
        if(tar_column==-1||tar_column==0)
            continue;
        QString file_name = ui->tableWidget_name->item(i,1)->text();
        QString dir_path = QCoreApplication::applicationDirPath();
        QDir tempDir;
        QString file_full_path = dir_path+"/"+file_name;
        if(!tempDir.exists(dir_path))
        {
            tempDir.mkpath(file_full_path);
        }
        QFile tempFile;
        tempDir.setCurrent(file_full_path);
        if(tempFile.exists(file_name))
        {
            qDebug()<<QString("file %1 has already exist").arg(file_name);
            bool b=tempFile.remove();
            int id=1;
            //            continue;
        }
        tempFile.setFileName(file_name);
        if(!tempFile.open(QIODevice::ReadWrite|QIODevice::Text|QIODevice::Truncate))
            return;
        qint64 pos = tempFile.size();
        tempFile.seek(pos);
        QAxObject * usedrange = worksheet->querySubObject("UsedRange");//获取该sheet的使用范围对象
        QVariant all_data = usedrange->property("Value");
        QVariantList all_list = all_data.toList();
        QString content_prefix;
        content_prefix = QString("[Info]\nLanguage=%1\n[String]\n").arg( ui->tableWidget_name->item(i,0)->text());
        tempFile.write(content_prefix.toUtf8());
        qDebug()<<QString("all_list.count() : %1").arg(all_list.count());
        //excel表格都是从1开始计数，第一行，第一列，但是数组列表都是从0开始，要注意
        if(default_key_column<0||default_key_column==0)
            default_key_column = 1;
        for(int i=0;i<all_list.count();++i)
        {
            QVariantList all_list_2 = all_list.at(i).toList();
            QString key_name = all_list_2.at(default_key_column-1).toString();
            if(key_name.isEmpty())
                continue;
            QString target_content = all_list_2.at(tar_column-1).toString();
            if(target_content.isEmpty())
                target_content = all_list_2.at(default_language_column-1).toString();
            if(target_content.isEmpty())
                continue;
            QString content = key_name+"="+target_content+"\n";
            tempFile.write(content.toUtf8());

        }
        tempFile.close();
    }
    workbook->dynamicCall("Close()");
}

bool MainWindow::CloseProcess(QString strExeName)
{
    strExeName.replace("/","\\");
    DWORD dwProcesses[1024], dwNeeded, dwProcNum;

    if ( !EnumProcesses( dwProcesses, sizeof(dwProcesses), &dwNeeded ) )
    {
        return false;
    }

    dwProcNum = dwNeeded / sizeof(DWORD);   /*计算进程数*/
    for ( int i = 0; i < dwProcNum; i++ )
    {
        HANDLE hProcess = OpenProcess( PROCESS_TERMINATE | PROCESS_VM_READ | PROCESS_QUERY_INFORMATION ,
                                       FALSE ,	dwProcesses[i] );
        if( NULL == hProcess )
        {
            continue;
        }

        HMODULE hMods[1024];
        DWORD	dwNeeded;
        if( EnumProcessModules(hProcess, hMods, sizeof(hMods), &dwNeeded))
        {
            char szModName[MAX_PATH];
            if ( GetModuleFileNameExA( hProcess, hMods[0], szModName,sizeof(szModName)))
            {
                QString TempName = QString::fromLocal8Bit(szModName);
                strExeName = strExeName.toUpper();
                TempName = TempName.toUpper();
                if (0 == QString::compare(strExeName,TempName, Qt::CaseInsensitive) )
                {
                    TerminateProcess(hProcess, 0);
                    return true;
                }
            }
        }
    }
    return false;
}

bool MainWindow::FindProcess(QString strExeName)
{
    strExeName.replace("/","\\");
    DWORD dwProcesses[1024], dwNeeded, dwProcNum;

    if (!EnumProcesses( dwProcesses, sizeof(dwProcesses), &dwNeeded ) )
    {
        return false;
    }

    dwProcNum = dwNeeded / sizeof(DWORD);   /*计算进程数*/
    for ( int i = 0; i < dwProcNum; i++ )
    {
        HANDLE hProcess = OpenProcess( PROCESS_TERMINATE | PROCESS_VM_READ | PROCESS_QUERY_INFORMATION ,
                                       FALSE ,	dwProcesses[i] );
        if( NULL == hProcess )
        {
            continue;
        }

        HMODULE hMods[1024];
        DWORD	dwNeeded;
        if( EnumProcessModules(hProcess, hMods, sizeof(hMods), &dwNeeded))
        {
            char szModName[MAX_PATH];
            if ( GetModuleFileNameExA( hProcess, hMods[0], szModName,sizeof(szModName)))
            {
                QString TempName = QString::fromLocal8Bit(szModName);
                strExeName = strExeName.toUpper();
                TempName = TempName.toUpper();
                if (0 == QString::compare(strExeName,TempName) )
                {
                    return true;
                }
            }
        }
    }

    return false;
}

QAxObject *MainWindow::read_excel(QString file_path)
{

    if(file_path.isEmpty())
        return nullptr;
    QList<QVariant> all_row_data;
    all_row_data.clear();
    QAxObject *excel = new QAxObject("Excel.Application");//连接Excel控件
    QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
    workbooks->dynamicCall("Open (const QString&)",file_path);//打开工作簿
    QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
    return workbook;
}

QAxObject *MainWindow::get_tar_work_sheet(QAxObject *workbook, QString worksheet_name)
{
    if(!workbook || worksheet_name.isEmpty())
        return nullptr;
    QAxObject *worksheets = workbook->querySubObject("WorkSheets");//获取所有的工作簿表
    if(!workbook)
        return nullptr;
    
    int worksheets_count = worksheets->property("Count").toInt();
    for(int i=0;i<worksheets_count;++i)
    {
        QAxObject *worksheet = worksheets->querySubObject("Item(int)",i+1);
        if(!worksheet)
            continue;
        QString sheet_name = worksheet->property("Name").toString();
        if(worksheet_name==sheet_name)
            return worksheet;
    }
    return nullptr;
}

int MainWindow::get_tar_sheet_column(QAxObject *worksheet, QString language_name)
{
    if(!worksheet)
        return -1;
    QAxObject * usedrange = worksheet->querySubObject("UsedRange");//获取该sheet的使用范围对象
    //    QAxObject* rows = usedrange->querySubObject("Rows");
    //    int nRows=rows->property("Count").toInt();
    QAxObject* columns=usedrange->querySubObject("Columns");
    int nColumns=columns->property("Count").toInt();
    for(int i=1;i<nColumns;++i)
    {
        QAxObject *range;
        range = worksheet->querySubObject("Cells(int,int)",1,i);
        if(!range)
            continue;
        QString str = range->dynamicCall("Value2()").toString();
        if(range->dynamicCall("Value2()").toString()==language_name)
            return i;
    }
    return -1;
}

void MainWindow::save_excel()
{
    QString filepath=QFileDialog::getSaveFileName(NULL,QObject::tr("Save orbit"),"/untitled.xlsx",QObject::tr("*.xlsx"));//获取保存路径
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
                aRowData.append(QVariant(/*row*column*/"haha"));
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

void MainWindow::set_connect()
{
    connect(ui->pushButton_open,&QPushButton::clicked,
            this,&MainWindow::on_open_file);
    connect(ui->pushButton_load,&QPushButton::clicked,
            this,&MainWindow::load_ini);
    connect(ui->pushButton_create,&QPushButton::clicked,
            this,&MainWindow::create_lang_morefast);
}
