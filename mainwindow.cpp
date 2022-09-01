#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QFileInfo>
#include <QMessageBox>


MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    initUI();

    m_recentFilePath = "D:/";
}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::initUI()
{
    connect(ui->btnBrowser1, SIGNAL(clicked()), this, SLOT(onBtnBrowserClicked1()));
    connect(ui->btnBrowser2, SIGNAL(clicked()), this, SLOT(onBtnBrowserClicked2()));
    connect(ui->btnBrowser3, SIGNAL(clicked()), this, SLOT(onBtnBrowserClicked3()));
}


QString MainWindow::browseExcelFilePath(const QString &info)
{
    QString filePath = QFileDialog::getOpenFileName(this, info, m_recentFilePath, "Excel files(*.xlsx *.xls");
    if(filePath.isEmpty())
    {
        return "";
    }

    QFileInfo fileInfo(filePath);
    m_recentFilePath = fileInfo.absolutePath();

    return filePath;
}


void MainWindow::onBtnBrowserClicked1()
{
    QString filePath = browseExcelFilePath(QString::fromLocal8Bit("选择表1 社保公积金结算单"));
    ui->editExcelFilePath1->setText(filePath);
}


void MainWindow::onBtnBrowserClicked2()
{
    QString filePath = browseExcelFilePath(QString::fromLocal8Bit("选择表2 薪资明细表（从系统中导出的）"));
    ui->editExcelFilePath2->setText(filePath);
}


void MainWindow::onBtnBrowserClicked3()
{
    QString filePath = browseExcelFilePath(QString::fromLocal8Bit("选择表3 薪资明细表（造册）"));
    ui->editExcelFilePath3->setText(filePath);
}


void MainWindow::onBtnCheck1()
{
    QString filePath1 = ui->editExcelFilePath1->text().trimmed();
    QString filePath2 = ui->editExcelFilePath2->text().trimmed();

    if(false == checkFileExists(filePath1)
    || false == checkFileExists(filePath2))
    {
        return;
    }



}


void MainWindow::onBtnCheck2()
{

}


void MainWindow::onBtnCheck3()
{


}

void MainWindow::showInfoMessageBox(const QString &info)
{
    QMessageBox::information(this, QString::fromLocal8Bit("提示信息"), info, QMessageBox::Yes);
}


 bool MainWindow::checkFileExists(const QString &filePath)
 {
     if(!QFileInfo::exists(filePath))
     {
         showInfoMessageBox(QString::fromLocal8Bit(QString("Not exist file: %1").arg(filePath).toStdString().c_str()));
         return false;
     }

     return true;
 }
