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
    m_excelVerify = new ExcelVerify();
}

MainWindow::~MainWindow()
{
    delete ui;
    delete m_excelVerify;
}


void MainWindow::initUI()
{
    ui->tabWidget->setCurrentIndex(0);

    connect(ui->btnBrowser1, SIGNAL(clicked()), this, SLOT(onBtnBrowserClicked1()));
    connect(ui->btnBrowser2, SIGNAL(clicked()), this, SLOT(onBtnBrowserClicked2()));
    connect(ui->btnBrowser3, SIGNAL(clicked()), this, SLOT(onBtnBrowserClicked3()));
    connect(ui->btnCheck1, SIGNAL(clicked()), this, SLOT(onBtnCheck1()));
    connect(ui->btnCheck2, SIGNAL(clicked()), this, SLOT(onBtnCheck2()));
    connect(ui->btnCheck3, SIGNAL(clicked()), this, SLOT(onBtnCheck3()));
}


QString MainWindow::browseExcelFilePath(const QString &info)
{
    QString filePath = QFileDialog::getOpenFileName(this, info, m_recentFilePath, "Excel files (*.xls *.xlsx)");
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

    m_excelVerify->checkExcelFile12(filePath1, filePath2);

}


void MainWindow::onBtnCheck2()
{
    QString filePath2 = ui->editExcelFilePath2->text().trimmed();
    QString filePath3 = ui->editExcelFilePath3->text().trimmed();

    if(false == checkFileExists(filePath2)
    || false == checkFileExists(filePath3))
    {
        return;
    }

    m_excelVerify->checkExcelFile23(filePath2, filePath3);
}


void MainWindow::onBtnCheck3()
{


}

void MainWindow::showInfoMessageBox(const QString &info)
{
    QMessageBox::information(this, "提示信息", info, QMessageBox::Yes);
}


 bool MainWindow::checkFileExists(const QString &filePath)
 {
     if(filePath.isEmpty())
     {
         showInfoMessageBox("文件路径不能为空");
         qWarning("empty file path");
         return false;
     }

     if(!QFileInfo::exists(filePath))
     {
         showInfoMessageBox(QString("错误的文件路径: %1").arg(filePath));
         qWarning("not vaild file path:%s", filePath.toLocal8Bit().data());
         return false;
     }

     return true;
 }
