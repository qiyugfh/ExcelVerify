#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QFileInfo>
#include <QMessageBox>
#include <QStandardItemModel>
#include <QStringList>
#include <QRegExpValidator>


const QString OK_TEXT_STYLE = "color:green";
const QString ERR_TEXT_STYLE = "color:red";

const QString CHECK_IN_INFO = "正在检查";
const QString CHECK_FINISH_INFO = "开始检查";

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


void MainWindow::resetTab3FormLabel()
{
    ui->lblGongJiJinGongSiJiaoNa->setStyleSheet(OK_TEXT_STYLE);
    ui->lblGongJiJinYuanGongJiaoNa->setStyleSheet(OK_TEXT_STYLE);
    ui->lblShiYeBaoXianGongSiJiaoNa->setStyleSheet(OK_TEXT_STYLE);
    ui->lblYiLiaoBaoXianGongSiJiaoNa->setStyleSheet(OK_TEXT_STYLE);
    ui->lblShiYeBaoXianYuanGongJiaoNa->setStyleSheet(OK_TEXT_STYLE);
    ui->lblYangLaoBaoXianGongSiJiaoNa->setStyleSheet(OK_TEXT_STYLE);
    ui->lblDaBingBaoXianYuanGongJiaoNa->setStyleSheet(OK_TEXT_STYLE);
    ui->lblYiLiaoBaoXianYuanGongJiaoNa->setStyleSheet(OK_TEXT_STYLE);
    ui->lblGongShangBaoXianGongSiJiaoNa->setStyleSheet(OK_TEXT_STYLE);
    ui->lblYangLaoBaoXianYuanGongJiaoNa->setStyleSheet(OK_TEXT_STYLE);
    ui->lblShengYuBaoXianGongSiJiaoNa->setStyleSheet(OK_TEXT_STYLE);

    ui->lblGongJiJinGongSiJiaoNa->setText("");
    ui->lblGongJiJinYuanGongJiaoNa->setText("");
    ui->lblShiYeBaoXianGongSiJiaoNa->setText("");
    ui->lblYiLiaoBaoXianGongSiJiaoNa->setText("");
    ui->lblShiYeBaoXianYuanGongJiaoNa->setText("");
    ui->lblYangLaoBaoXianGongSiJiaoNa->setText("");
    ui->lblDaBingBaoXianYuanGongJiaoNa->setText("");
    ui->lblYiLiaoBaoXianYuanGongJiaoNa->setText("");
    ui->lblGongShangBaoXianGongSiJiaoNa->setText("");
    ui->lblYangLaoBaoXianYuanGongJiaoNa->setText("");
    ui->lblShengYuBaoXianGongSiJiaoNa->setText("");
}


void MainWindow::initTab3Form()
{
    //非负浮点数
    QRegExpValidator *regValidator = new QRegExpValidator(QRegExp("^\\d+(\\.\\d+)?$"), this);
    ui->leditGongJiJinGongSiJiaoNa->setValidator(regValidator);
    ui->leditGongJiJinYuanGongJiaoNa->setValidator(regValidator);
    ui->leditShiYeBaoXianGongSiJiaoNa->setValidator(regValidator);
    ui->leditYiLiaoBaoXianGongSiJiaoNa->setValidator(regValidator);
    ui->leditShengYuBaoXianGongSiJiaoNa->setValidator(regValidator);
    ui->leditShiYeBaoXianYuanGongJiaoNa->setValidator(regValidator);
    ui->leditYangLaoBaoXianGongSiJiaoNA->setValidator(regValidator);
    ui->leditDaBingBaoXianYuanGongJiaoNa->setValidator(regValidator);
    ui->leditYiLiaoBaoXianYuanGongJiaoNa->setValidator(regValidator);
    ui->leditGongShangBaoXianGongSiJiaoNa->setValidator(regValidator);
    ui->leditYangLaoBaoXianYuanGongJiaoNa->setValidator(regValidator);

    resetTab3FormLabel();
}


void MainWindow::initTableWidgetStyle(QTableWidget *tw, const QStringList &labels)
{
    tw->setColumnCount(labels.size());
    tw->setRowCount(0);
    tw->setHorizontalHeaderLabels(labels);

    //设置表头样式
    tw->horizontalHeader()->setStyleSheet("QHeaderView::section{background:rgb(245,245,245);}");
    //设置纵向表头
    tw->verticalHeader()->setStyleSheet("QHeaderView::section{background:rgb(245,245,245);}");
    //设置为只读模式
    tw->setEditTriggers(QAbstractItemView::NoEditTriggers);
}


void MainWindow::initTableWidget1()
{
    QStringList labels;
    labels << "工号" << "姓名" << "单位养老金" << "单位医保" << "单位失业金" << "单位工伤保险" << "单位生育保险"
            << "个人养老金" << "个人医保" << "个人失业金" << "个人大病医疗" << "单位公积金" << "个人公积金";

    initTableWidgetStyle(this->ui->tableWidget1, labels);
}


void MainWindow::initTableWidget2()
{
    QStringList labels;
    labels << "工号" << "姓名" << "夜班津贴" << "餐贴" << "其他奖金/资金归集专项激励" << "通讯补贴" << "电脑补贴" << "加班补贴"
           << "佣金" << "税前其他补发" << "税前其他补扣" /*<< "养老保险公司缴纳" << "医疗保险公司缴纳" << "生育保险公司缴纳"
           << "失业保险公司缴纳" << "工伤保险公司缴纳" << "公积金公司缴纳" << "养老保险员工缴纳" << "医疗保险员工缴纳"
           <<"失业保险员工缴纳" << "大病保险员工缴纳" << "公积金员工缴纳"*/;

    initTableWidgetStyle(this->ui->tableWidget2, labels);
}


void MainWindow::initUI()
{
    this->setWindowTitle("社保公积金核算工具");
    ui->tabWidget->setCurrentIndex(0);

    initTableWidget1();
    initTableWidget2();
    initTab3Form();

    connect(ui->btnBrowser1, SIGNAL(clicked()), this, SLOT(onBtnBrowserClicked1()));
    connect(ui->btnBrowser2, SIGNAL(clicked()), this, SLOT(onBtnBrowserClicked2()));
    connect(ui->btnBrowser3, SIGNAL(clicked()), this, SLOT(onBtnBrowserClicked3()));
    connect(ui->btnCheck1, SIGNAL(clicked()), this, SLOT(onBtnCheck1()));
    connect(ui->btnCheck2, SIGNAL(clicked()), this, SLOT(onBtnCheck2()));
    connect(ui->btnCheck3, SIGNAL(clicked()), this, SLOT(onBtnCheck3()));
}


QString MainWindow::browseExcelFilePath(const QString &info)
{
    QString filePath = QFileDialog::getOpenFileName(this, info, m_recentFilePath, "Excel files (*.xlsx)");
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

void MainWindow::delDataFromTableWidget(QTableWidget *tw)
{
    tw->setRowCount(0);
}


void MainWindow::insertTableWidgetItem(QTableWidget *tw, int row, int col, const QString &info, bool error)
{
    QTableWidgetItem *item = NULL;
    item = new QTableWidgetItem();
    if(error)
    {
        item->setBackground(QColor("yellow"));
    }

    QFont font;
    font.setPixelSize(24);
    item->setFont(font);

    item->setTextAlignment(Qt::AlignVCenter);
    item->setText(info);
    tw->setItem(row, col, item);
}


void MainWindow::addDataToTableWidget1()
{
    QTableWidget *tw = ui->tableWidget1;
    int rowCount = m_excelVerify->m_checkResultList1.size();

    showInfoMessageBox(QString("检查结束，共%1人信息错误").arg(QString::number(rowCount)));
    tw->setRowCount(rowCount);

    for(int i = 0; i < rowCount; i++)
    {
        int col = 0;
        ExcelFileCheckErrorRowStruct &rowStruct = m_excelVerify->m_checkResultList1[i];
        insertTableWidgetItem(tw, i, col++, rowStruct.gongHao);
        insertTableWidgetItem(tw, i, col++, rowStruct.xingMing);
        for(int j = 0; j < rowStruct.checkCellResults.size(); j++)
        {
            QString info = QString::number(rowStruct.checkCellResults[j].val, 'f', 2);
            bool error = rowStruct.checkCellResults[j].error;
            insertTableWidgetItem(tw, i, col++, info, error);
        }
    }
}


void MainWindow::addDataToTableWidget2()
{
    QTableWidget *tw = ui->tableWidget2;
    int rowCount = m_excelVerify->m_checkResultList2.size();

    showInfoMessageBox(QString("检查结束，共%1人信息错误").arg(QString::number(rowCount)));
    tw->setRowCount(rowCount);

    for(int i = 0; i < rowCount; i++)
    {
        int col = 0;
        ExcelFileCheckErrorRowStruct &rowStruct = m_excelVerify->m_checkResultList2[i];
        insertTableWidgetItem(tw, i, col++, rowStruct.gongHao);
        insertTableWidgetItem(tw, i, col++, rowStruct.xingMing);
        for(int j = 0; j < rowStruct.checkCellResults.size(); j++)
        {
            QString info = QString::number(rowStruct.checkCellResults[j].val, 'f', 2);
            bool error = rowStruct.checkCellResults[j].error;
            insertTableWidgetItem(tw, i, col++, info, error);
        }
    }
}


void MainWindow::onBtnCheck1()
{
    delDataFromTableWidget(ui->tableWidget1);

    QString filePath1 = ui->editExcelFilePath1->text().trimmed();
    QString filePath2 = ui->editExcelFilePath2->text().trimmed();

    if(false == checkFileExists(filePath1)
    || false == checkFileExists(filePath2))
    {
        return;
    }

    ui->btnCheck1->setText(CHECK_IN_INFO);
    m_excelVerify->checkExcelFile12(filePath1, filePath2);

    addDataToTableWidget1();

    ui->btnCheck1->setText(CHECK_FINISH_INFO);

}


void MainWindow::onBtnCheck2()
{
    delDataFromTableWidget(ui->tableWidget2);

    QString filePath2 = ui->editExcelFilePath2->text().trimmed();
    QString filePath3 = ui->editExcelFilePath3->text().trimmed();

    if(false == checkFileExists(filePath2)
    || false == checkFileExists(filePath3))
    {
        return;
    }

    ui->btnCheck2->setText(CHECK_IN_INFO);
    m_excelVerify->checkExcelFile23(filePath2, filePath3);

    addDataToTableWidget2();

    ui->btnCheck2->setText(CHECK_FINISH_INFO);
}


void MainWindow::compareTab3FormInput()
{
    float inputVal = 0.0f;
    float realVal = 0.0f;

    inputVal = ui->leditGongJiJinGongSiJiaoNa->text().toFloat();
    realVal = m_excelVerify->m_excelColSum.gongJiJinBaoXianGongSiJiaoNa;
    ui->lblGongJiJinGongSiJiaoNa->setText(QString::number(realVal, 'f', 2));
    if(abs(inputVal - realVal) >= CHECK_FLOAT_PRECISION)
    {
        ui->lblGongJiJinGongSiJiaoNa->setStyleSheet(ERR_TEXT_STYLE);
    }

    inputVal = ui->leditGongJiJinYuanGongJiaoNa->text().toFloat();
    realVal = m_excelVerify->m_excelColSum.gongJiJinYuanGongJiaoNa;
    ui->lblGongJiJinYuanGongJiaoNa->setText(QString::number(realVal, 'f', 2));
    if(abs(inputVal - realVal) >= CHECK_FLOAT_PRECISION)
    {
        ui->lblGongJiJinYuanGongJiaoNa->setStyleSheet(ERR_TEXT_STYLE);
    }

    inputVal = ui->leditShiYeBaoXianGongSiJiaoNa->text().toFloat();
    realVal = m_excelVerify->m_excelColSum.shiYeBaoXianGongSiJiaoNa;
    ui->lblShiYeBaoXianGongSiJiaoNa->setText(QString::number(realVal, 'f', 2));
    if(abs(inputVal - realVal) >= CHECK_FLOAT_PRECISION)
    {
        ui->lblShiYeBaoXianGongSiJiaoNa->setStyleSheet(ERR_TEXT_STYLE);
    }

    inputVal = ui->leditYiLiaoBaoXianGongSiJiaoNa->text().toFloat();
    realVal = m_excelVerify->m_excelColSum.yiLiaoBaoXianGongSiJiaoNa;
    ui->lblYiLiaoBaoXianGongSiJiaoNa->setText(QString::number(realVal, 'f', 2));
    if(abs(inputVal - realVal) >= CHECK_FLOAT_PRECISION)
    {
        ui->lblYiLiaoBaoXianGongSiJiaoNa->setStyleSheet(ERR_TEXT_STYLE);
    }

    inputVal = ui->leditShiYeBaoXianYuanGongJiaoNa->text().toFloat();
    realVal = m_excelVerify->m_excelColSum.shiYeBaoXianYuanGongJiaoNa;
    ui->lblShiYeBaoXianYuanGongJiaoNa->setText(QString::number(realVal, 'f', 2));
    if(abs(inputVal - realVal) >= CHECK_FLOAT_PRECISION)
    {
        ui->lblShiYeBaoXianYuanGongJiaoNa->setStyleSheet(ERR_TEXT_STYLE);
    }

    inputVal = ui->leditYangLaoBaoXianGongSiJiaoNA->text().toFloat();
    realVal = m_excelVerify->m_excelColSum.yangLaoBaoXianGongSiJiaoNa;
    ui->lblYangLaoBaoXianGongSiJiaoNa->setText(QString::number(realVal, 'f', 2));
    if(abs(inputVal - realVal) >= CHECK_FLOAT_PRECISION)
    {
        ui->lblYangLaoBaoXianGongSiJiaoNa->setStyleSheet(ERR_TEXT_STYLE);
    }

    inputVal = ui->leditDaBingBaoXianYuanGongJiaoNa->text().toFloat();
    realVal = m_excelVerify->m_excelColSum.daBingBaoXianYuanGongJiaoNa;
    ui->lblDaBingBaoXianYuanGongJiaoNa->setText(QString::number(realVal, 'f', 2));
    if(abs(inputVal - realVal) >= CHECK_FLOAT_PRECISION)
    {
        ui->lblDaBingBaoXianYuanGongJiaoNa->setStyleSheet(ERR_TEXT_STYLE);
    }

    inputVal = ui->leditYiLiaoBaoXianYuanGongJiaoNa->text().toFloat();
    realVal = m_excelVerify->m_excelColSum.yiLiaoBaoXianYuanGongJiaoNa;
    ui->lblYiLiaoBaoXianYuanGongJiaoNa->setText(QString::number(realVal, 'f', 2));
    if(abs(inputVal - realVal) >= CHECK_FLOAT_PRECISION)
    {
        ui->lblYiLiaoBaoXianYuanGongJiaoNa->setStyleSheet(ERR_TEXT_STYLE);
    }

    inputVal = ui->leditGongShangBaoXianGongSiJiaoNa->text().toFloat();
    realVal = m_excelVerify->m_excelColSum.gongShangBaoXianGongSiJiaoNa;
    ui->lblGongShangBaoXianGongSiJiaoNa->setText(QString::number(realVal, 'f', 2));
    if(abs(inputVal - realVal) >= CHECK_FLOAT_PRECISION)
    {
        ui->lblGongShangBaoXianGongSiJiaoNa->setStyleSheet(ERR_TEXT_STYLE);
    }

    inputVal = ui->leditYangLaoBaoXianYuanGongJiaoNa->text().toFloat();
    realVal = m_excelVerify->m_excelColSum.yangLaoBaoXianYuanGongJiaoNa;
    ui->lblYangLaoBaoXianYuanGongJiaoNa->setText(QString::number(realVal, 'f', 2));
    if(abs(inputVal - realVal) >= CHECK_FLOAT_PRECISION)
    {
        ui->lblYangLaoBaoXianYuanGongJiaoNa->setStyleSheet(ERR_TEXT_STYLE);
    }

    inputVal = ui->leditShengYuBaoXianGongSiJiaoNa->text().toFloat();
    realVal = m_excelVerify->m_excelColSum.shengYuBaoXianGongSiJiaoNa;
    ui->lblShengYuBaoXianGongSiJiaoNa->setText(QString::number(realVal, 'f', 2));
    if(abs(inputVal - realVal) >= CHECK_FLOAT_PRECISION)
    {
        ui->lblShengYuBaoXianGongSiJiaoNa->setStyleSheet(ERR_TEXT_STYLE);
    }
}


void MainWindow::onBtnCheck3()
{
    resetTab3FormLabel();

    QString filePath1 = ui->editExcelFilePath1->text().trimmed();
    if(false == checkFileExists(filePath1))
    {
        return;
    }

    m_excelVerify->checkExcelFile110(filePath1);
    compareTab3FormInput();
}

void MainWindow::showInfoMessageBox(const QString &info)
{
    QMessageBox::information(this, "提示信息", info, "确认");
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
