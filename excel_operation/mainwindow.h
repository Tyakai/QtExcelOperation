#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QAxObject>
#include <QAxWidget>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_open_file(bool);
    void load_ini();
    void create_lang();
    void create_lang_morefast();

private:
    bool CloseProcess(QString strExeName);
    bool FindProcess(QString strExeName);
    QAxObject* read_excel(QString file_path);
    QAxObject* get_tar_work_sheet(QAxObject* workbook,QString worksheet_name);
    int        get_tar_sheet_column(QAxObject* worksheet,QString language_name);
    QString ToUnicode(const QString& cstr);
    int        default_language_column=-1;
    void save_excel();
    void set_connect();
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
