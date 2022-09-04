#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QTableWidget>
#include "excelverify.h"


QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void onBtnBrowserClicked1();
    void onBtnBrowserClicked2();
    void onBtnBrowserClicked3();
    void onBtnCheck1();
    void onBtnCheck2();
    void onBtnCheck3();



private:
    void initUI();
    QString browseExcelFilePath(const QString &info);
    void showInfoMessageBox(const QString &info);
    bool checkFileExists(const QString &filePath);
    void initTableWidget1();
    void initTableWidget2();
    void initTableWidgetStyle(QTableWidget *tw, const QStringList &labels);
    void delDataFromTableWidget(QTableWidget *tw);
    void addDataToTableWidget1();
    void addDataToTableWidget2();
    void insertTableWidgetItem(QTableWidget *tw, int row, int col, const QString &info, bool error = false);
    void initTab3Form();
    void resetTab3FormLabel();
    void compareTab3FormInput();


private:
    Ui::MainWindow *ui;
    QString         m_recentFilePath;
    QString         m_excelFile1;
    QString         m_excelFile2;
    QString         m_excelFile3;


    ExcelVerify     *m_excelVerify;
};
#endif // MAINWINDOW_H
