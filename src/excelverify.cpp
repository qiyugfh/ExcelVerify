#include "excelverify.h"
#include "log.h"

// QXlsx headers
#include "xlsxdocument.h"
#include "xlsxworksheet.h"
#include "xlsxcell.h"


ExcelVerify::ExcelVerify()
{
    initMap1();
    initMap2();
    initMap3();

}


ExcelVerify::~ExcelVerify()
{

}


void ExcelVerify::initMap1()
{
    // 第1列序号
    m_map1.insert("工号", 2);
    m_map1.insert("姓名", 3);
    m_map1.insert("单位养老金", 9);
    m_map1.insert("单位医保", 10);
    m_map1.insert("单位失业金", 11);
    m_map1.insert("单位工伤保险", 12);
    m_map1.insert("单位生育保险", 13);
    m_map1.insert("个人养老金", 14);
    m_map1.insert("个人医保", 15);
    m_map1.insert("个人失业金", 16);
    m_map1.insert("个人大病医疗", 17);
    m_map1.insert("单位公积金", 21);
    m_map1.insert("个人公积金", 22);
}


void ExcelVerify::initMap2()
{
    // 第1列工号
    m_map2.insert("工号", 1);
    m_map2.insert("姓名", 2);
    m_map2.insert("夜班津贴", 35);
    m_map2.insert("餐贴", 44);
    m_map2.insert("通讯补贴", 45);
    m_map2.insert("电脑补贴", 46);
    m_map2.insert("加班补贴", 47);
    m_map2.insert("其他奖金/资金归集专项激励", 69);
    m_map2.insert("税前其他补发", 71);
    m_map2.insert("佣金", 75);
    m_map2.insert("税前其他补扣", 80);

    m_map2.insert("养老保险公司缴纳", 93);
    m_map2.insert("医疗保险公司缴纳", 94);
    m_map2.insert("生育保险公司缴纳", 95);
    m_map2.insert("失业保险公司缴纳", 96);
    m_map2.insert("工伤保险公司缴纳", 97);
    m_map2.insert("公积金公司缴纳", 100);
    m_map2.insert("养老保险员工缴纳", 102);
    m_map2.insert("医疗保险员工缴纳", 103);
    m_map2.insert("失业保险员工缴纳", 105);
    m_map2.insert("大病保险员工缴纳", 108);
    m_map2.insert("公积金员工缴纳", 109);
}


void ExcelVerify::initMap3()
{
    // 第1列序号
    m_map3.insert("工号", 3);
    m_map3.insert("姓名", 2);
    m_map3.insert("通讯补贴", 12);
    m_map3.insert("电脑补贴", 13);
    m_map3.insert("佣金激励", 14);
    m_map3.insert("夜班补贴", 15);
    m_map3.insert("加班补贴", 16);
    m_map3.insert("其他补贴", 17);
    m_map3.insert("补发补扣", 18);
}





bool ExcelVerify::checkExcelFile12(const QString &excelFile1, const QString &excelFile2)
{

    if(false == readExcelFile1(excelFile1))
    {
        return false;
    }

    if(false == readExcelFile2(excelFile2))
    {
        return false;
    }



    return true;
}


/*表2（薪资明细表） 与表3（薪资明细表2.0）的关系
薪资明细表 AI（夜班津贴）=薪资明细表2.0 O（夜班补贴）
薪资明细表 AS（通讯补贴）=薪资明细表2.0 L（通讯补贴）
薪资明细表 AT（电脑补贴）=薪资明细表2.0 M（电脑补贴）
薪资明细表 AU（加班补贴）=薪资明细表2.0 P（加班补贴）
薪资明细表 BQ（其他奖金/资金归集专项激励）+薪资明细表 AR（餐贴）=薪资明细表2.0 Q（其他补贴）
薪资明细表 BW（佣金）=薪资明细表2.0 N（佣金激励）
薪资明细表 BS（税前其他补发）+薪资明细表 CB（税前其他补扣）=薪资明细表2.0 R（补发补扣）
*/
bool ExcelVerify::checkExcelFile23(const QString &excelFile2, const QString &excelFile3)
{


    return true;
}


bool ExcelVerify::readExcelFile1(const QString &excelFile)
{
    m_excelFileRowStructList1.clear();

    qInfo("begin load excel file: %s", excelFile.toStdString().c_str());

    QXlsx::Document xlsxDoc(excelFile);
    if(false == xlsxDoc.load())
    {
        qCritical("Failed to load excel %s", excelFile.toStdString().c_str());
        return false;
    }

    xlsxDoc.selectSheet(0);

    QXlsx::Worksheet *workSheet = xlsxDoc.currentWorksheet();
    if(NULL == workSheet)
    {
        qCritical("work sheet is NULL");
        return false;
    }

    // 获取有效范围
    QXlsx::CellRange cellRange = workSheet->dimension();
    int rowCount = cellRange.rowCount();
    int colCount = cellRange.columnCount();

    // 表1 第2、2行是表头
    for(int i = 3; i <= rowCount; i++)
    {
        ExcelFile1RowStruct rowStruct;
        for(int j = 1; j <= colCount; j++)
        {
            QXlsx::Cell *cell = workSheet->cellAt(i, j);
            if(NULL == cell)
            {
                qWarning("cell(%d, %d) is null", i, j);
                continue;
            }

            switch(j)
            {
            case 2:
                rowStruct.gongHao = cell->value().toString();
                break;
            case 3:
                rowStruct.xingMing = cell->value().toString();
                break;
            case 9:
                rowStruct.danWeiYangLaoJin = cell->value().toFloat();
                break;
            case 10:
                rowStruct.danWeiYiBao = cell->value().toFloat();
                break;
            case 11:
                rowStruct.danWeiShiYeJin = cell->value().toFloat();
                break;
            case 12:
                rowStruct.danWeiGongShangBaoXian = cell->value().toFloat();
                break;
            case 13:
                rowStruct.danWeiShengYuBaoXian = cell->value().toFloat();
                break;
            case 14:
                rowStruct.geRenYangLaoJin = cell->value().toFloat();
                break;
            case 15:
                rowStruct.geRenYiBao = cell->value().toFloat();
                break;
            case 16:
                rowStruct.geRenShiYeJin = cell->value().toFloat();
                break;
            case 17:
                rowStruct.geRenDaBingYiLiao = cell->value().toFloat();
                break;
            case 21:
                rowStruct.danWeiGongJiJin = cell->value().toFloat();
                break;
            case 22:
                rowStruct.geRenGongJiJin = cell->value().toFloat();
                break;
            default:
                break;
            }
        }

        m_excelFileRowStructList1.push_back(rowStruct);
    }

    if(m_excelFileRowStructList1.isEmpty())
    {
        qWarning("excel file no data");
        return false;
    }

    return true;
}

bool ExcelVerify::readExcelFile2(const QString &excelFile)
{
    m_excelFileRowStructList2.clear();

    qInfo("begin load excel file: %s", excelFile.toStdString().c_str());

    QXlsx::Document xlsxDoc(excelFile);
    if(false == xlsxDoc.load())
    {
        qCritical("Failed to load excel %s", excelFile.toStdString().c_str());
        return false;
    }

    xlsxDoc.selectSheet(0);

    QXlsx::Worksheet *workSheet = xlsxDoc.currentWorksheet();
    if(NULL == workSheet)
    {
        qCritical("work sheet is NULL");
        return false;
    }

    int rowCount = workSheet->dimension().rowCount();
    int colCount = workSheet->dimension().columnCount();

    // 表2 第1行是表头
    for(int i = 2; i <= rowCount; i++)
    {
        ExcelFile2RowStruct rowStruct;
        for(int j = 1; j <= colCount; j++)
        {
            QXlsx::Cell *cell = workSheet->cellAt(i, j);
            if(NULL == cell)
            {
                qWarning("cell(%d, %d) is null", i, j);
                continue;
            }

            switch(j)
            {
            case 1:
                rowStruct.gongHao = cell->value().toString();
                break;
            case 2:
                rowStruct.xingMing = cell->value().toString();
                break;
            case 35:
                rowStruct.yeBanJinTie = cell->value().toFloat();
                break;
            case 44:
                rowStruct.canTie = cell->value().toFloat();
                break;
            case 45:
                rowStruct.tongXunBuTie = cell->value().toFloat();
                break;
            case 46:
                rowStruct.dianNaoBuTie = cell->value().toFloat();
                break;
            case 47:
                rowStruct.jiaBanBuTie = cell->value().toFloat();
                break;
            case 69:
                rowStruct.qiTaJiangJinZiJinGuiJiZhuanXiangJiLi = cell->value().toFloat();
                break;
            case 71:
                rowStruct.shuiQianQiTaBuFa = cell->value().toFloat();
                break;
            case 75:
                rowStruct.yongJin = cell->value().toFloat();
                break;
            case 80:
                rowStruct.shuiQianQiTaBuKou = cell->value().toFloat();
                break;
            case 93:
                rowStruct.yangLaoBaoXianGongSiJiaoNa = cell->value().toFloat();
                break;
            case 94:
                rowStruct.yiLiaoBaoXianGongSiJiaoNa = cell->value().toFloat();
                break;
            case 95:
                rowStruct.shengYuBaoXianGongSiJiaoNa = cell->value().toFloat();
                break;
            case 96:
                rowStruct.shiYeBaoXianGongSiJiaoNa = cell->value().toFloat();
                break;
            case 97:
                rowStruct.gongShangBaoXianGongSiJiaoNa = cell->value().toFloat();
                break;
            case 100:
                rowStruct.gongJiJinGongSiJiaoNa = cell->value().toFloat();
                break;
            case 102:
                rowStruct.yangLaoBaoXianYuanGongJiaoNa = cell->value().toFloat();
                break;
            case 103:
                rowStruct.yiLiaoBaoXianYuanGongJiaoNa = cell->value().toFloat();
                break;
            case 105:
                rowStruct.shiYeBaoXianYuanGongJiaoNa = cell->value().toFloat();
                break;
            case 108:
                rowStruct.daBingBaoXianYuanGongJiaoNa = cell->value().toFloat();
                break;
            case 109:
                rowStruct.gongJiJinYuanGongJiaoNa = cell->value().toFloat();
                break;
            default:
                break;
            }
        }

        m_excelFileRowStructList2.push_back(rowStruct);
    }

    if(m_excelFileRowStructList2.isEmpty())
    {
        qWarning("excel file no data");
        return false;
    }

    return true;
}


bool ExcelVerify::readExcelFile3(const QString &excelFile)
{
    return true;
}



