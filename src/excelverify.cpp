#include "excelverify.h"
#include "log.h"


// QXlsx headers
#include "xlsxdocument.h"
#include "xlsxworksheet.h"
#include "xlsxcell.h"


ExcelVerify::ExcelVerify()
{

}


ExcelVerify::~ExcelVerify()
{

}


// 相同值返回false，不同值返回true
bool ExcelVerify::checkFloatValueIsDiff(float val1, float val2)
{
    if(qAbs(val1 - val2) <= CHECK_FLOAT_PRECISION)
    {
        return false;
    }

    return true;
}


bool ExcelVerify::writeCheckResult1ToFile()
{
    return true;
}


bool ExcelVerify::checkExcelFile12(const QString &excelFile1, const QString &excelFile2)
{
    m_checkResultList1.clear();

    if(false == readExcelFile1(excelFile1))
    {
        return false;
    }

    if(false == readExcelFile2(excelFile2))
    {
        return false;
    }

    for(int i = 0; i < m_excelFileRowStructList1.size(); i++)
    {
        const ExcelFile1RowStruct &s1 = m_excelFileRowStructList1.at(i);

        for(int j = 0; j < m_excelFileRowStructList2.size(); j++)
        {
            const ExcelFile2RowStruct &s2 = m_excelFileRowStructList2.at(j);
            if(0 != s1.gongHao.compare(s2.gongHao))
            {
                continue;
            }

            ExcelFileCheckErrorRowStruct rowResult;
            rowResult.gongHao = s1.gongHao;
            rowResult.xingMing = s1.xingMing;

            CheckCellResult cellResult;

            cellResult.error = checkFloatValueIsDiff(s1.danWeiYangLaoJin, s2.yangLaoBaoXianGongSiJiaoNa);
            cellResult.val = s1.danWeiYangLaoJin;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s1.danWeiYiBao, s2.yiLiaoBaoXianGongSiJiaoNa);
            cellResult.val = s1.danWeiYiBao;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s1.danWeiShiYeJin, s2.shiYeBaoXianGongSiJiaoNa);
            cellResult.val = s1.danWeiShiYeJin;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s1.danWeiGongShangBaoXian, s2.gongShangBaoXianGongSiJiaoNa);
            cellResult.val = s1.danWeiGongShangBaoXian;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s1.danWeiShengYuBaoXian, s2.shengYuBaoXianGongSiJiaoNa);
            cellResult.val = s1.danWeiShengYuBaoXian;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s1.geRenYangLaoJin, s2.yangLaoBaoXianYuanGongJiaoNa);
            cellResult.val = s1.geRenYangLaoJin;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s1.geRenYiBao, s2.yiLiaoBaoXianYuanGongJiaoNa);
            cellResult.val = s1.geRenYiBao;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s1.geRenShiYeJin, s2.shiYeBaoXianYuanGongJiaoNa);
            cellResult.val = s1.geRenShiYeJin;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s1.geRenDaBingYiLiao, s2.daBingBaoXianYuanGongJiaoNa);
            cellResult.val = s1.geRenDaBingYiLiao;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s1.danWeiGongJiJin, s2.gongJiJinGongSiJiaoNa);
            cellResult.val = s1.danWeiGongJiJin;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s1.geRenGongJiJin, s2.gongJiJinYuanGongJiaoNa);
            cellResult.val = s1.geRenGongJiJin;
            rowResult.checkCellResults.push_back(cellResult);

            bool hasError = false;
            for(int k = 0; k < rowResult.checkCellResults.size(); k++)
            {
                if(rowResult.checkCellResults[k].error)
                {
                    hasError = true;
                    break;
                }
            }

            if(hasError)
            {
                m_checkResultList1.push_back(rowResult);
            }
        }
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
    m_checkResultList2.clear();

    if(false == readExcelFile2(excelFile2))
    {
        return false;
    }

    if(false == readExcelFile3(excelFile3))
    {
        return false;
    }

    for(int i = 0; i < m_excelFileRowStructList2.size(); i++)
    {
        const ExcelFile2RowStruct &s2 = m_excelFileRowStructList2.at(i);

        for(int j = 0; j < m_excelFileRowStructList3.size(); j++)
        {
            const ExcelFile3RowStruct &s3 = m_excelFileRowStructList3.at(j);
            if(0 != s2.gongHao.compare(s3.gongHao))
            {
                continue;
            }

            ExcelFileCheckErrorRowStruct rowResult;
            rowResult.gongHao = s2.gongHao;
            rowResult.xingMing = s2.xingMing;

            CheckCellResult cellResult;

            cellResult.error = checkFloatValueIsDiff(s2.yeBanJinTie, s3.yeBanBuTie);
            cellResult.val = s2.yeBanJinTie;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s2.canTie + s2.qiTaJiangJinZiJinGuiJiZhuanXiangJiLi, s3.qiTaBuTie);
            cellResult.val = s2.canTie;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.val = s2.qiTaJiangJinZiJinGuiJiZhuanXiangJiLi;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s2.tongXunBuTie, s3.tongXunBuTie);
            cellResult.val = s2.tongXunBuTie;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s2.dianNaoBuTie, s3.dianNaoBuTie);
            cellResult.val = s2.dianNaoBuTie;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s2.jiaBanBuTie, s3.jiaBanBuTie);
            cellResult.val = s2.jiaBanBuTie;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s2.yongJin, s3.yongJinJiLi);
            cellResult.val = s2.yongJin;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.error = checkFloatValueIsDiff(s2.shuiQianQiTaBuFa + s2.shuiQianQiTaBuKou, s3.buFaBuKou);
            cellResult.val = s2.shuiQianQiTaBuFa;
            rowResult.checkCellResults.push_back(cellResult);

            cellResult.val = s2.shuiQianQiTaBuKou;
            rowResult.checkCellResults.push_back(cellResult);

            bool hasError = false;
            for(int k = 0; k < rowResult.checkCellResults.size(); k++)
            {
                if(rowResult.checkCellResults[k].error)
                {
                    hasError = true;
                    break;
                }
            }

            if(hasError)
            {
                m_checkResultList2.push_back(rowResult);
            }
        }
    }

    return true;
}



/*
表1 社保公积金结算单.xlsx
列名，列序号
"工号", 2
"姓名", 3
"单位养老金", 9
"单位医保", 10
"单位失业金", 11
"单位工伤保险", 12
"单位生育保险", 13
"个人养老金", 14
"个人医保", 15
"个人失业金", 16
"个人大病医疗", 17
"单位公积金", 21
"个人公积金", 22
*/
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

    // 表1 第1、2行是表头
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
                rowStruct.gongHao = cell->value().toString().trimmed();
                break;
            case 3:
                rowStruct.xingMing = cell->value().toString().trimmed();
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


/*
表2 薪资明细表.xlsx
列名，列序号
"工号", 1
"姓名", 2
"夜班津贴", 35
"餐贴", 44
"通讯补贴", 45
"电脑补贴", 46
"加班补贴", 47
"其他奖金/资金归集专项激励", 69
"税前其他补发", 71
"佣金", 75
"税前其他补扣", 80
"养老保险公司缴纳", 93
"医疗保险公司缴纳", 94
"生育保险公司缴纳", 95
"失业保险公司缴纳", 96
"工伤保险公司缴纳", 97
"公积金公司缴纳", 100
"养老保险员工缴纳", 102
"医疗保险员工缴纳", 103
"失业保险员工缴纳", 105
"大病保险员工缴纳", 108
"公积金员工缴纳", 109
*/
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

    xlsxDoc.selectSheet("薪资明细表");

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
                rowStruct.gongHao = cell->value().toString().trimmed();
                break;
            case 2:
                rowStruct.xingMing = cell->value().toString().trimmed();
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


/*
表3 薪资明细表2.0.xlsx
列名，列序号
"工号", 3
"姓名", 4
"通讯补贴", 12
"电脑补贴", 13
"佣金激励", 14
"夜班补贴", 15
"加班补贴", 16
"其他补贴", 17
"补发补扣", 18
*/
bool ExcelVerify::readExcelFile3(const QString &excelFile)
{
    m_excelFileRowStructList3.clear();

    qInfo("begin load excel file: %s", excelFile.toStdString().c_str());

    QXlsx::Document xlsxDoc(excelFile);
    if(false == xlsxDoc.load())
    {
        qCritical("Failed to load excel %s", excelFile.toStdString().c_str());
        return false;
    }

    // 表3 第1个sheet是“汇总”，是隐藏sheet
    xlsxDoc.selectSheet("明细");

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
        ExcelFile3RowStruct rowStruct;
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
            case 3:
                rowStruct.gongHao = cell->value().toString().trimmed();
                break;
            case 4:
                rowStruct.xingMing = cell->value().toString().trimmed();
                break;
            case 12:
                rowStruct.tongXunBuTie = cell->value().toFloat();
                break;
            case 13:
                rowStruct.dianNaoBuTie = cell->value().toFloat();
                break;
            case 14:
                rowStruct.yongJinJiLi = cell->value().toFloat();
                break;
            case 15:
                rowStruct.yeBanBuTie = cell->value().toFloat();
                break;
            case 16:
                rowStruct.jiaBanBuTie = cell->value().toFloat();
                break;
            case 17:
                rowStruct.qiTaBuTie = cell->value().toFloat();
                break;
            case 18:
                rowStruct.buFaBuKou = cell->value().toFloat();
                break;
            default:
                break;
            }
        }

        m_excelFileRowStructList3.push_back(rowStruct);
    }

    if(m_excelFileRowStructList3.isEmpty())
    {
        qWarning("excel file no data");
        return false;
    }

    return true;
}


bool ExcelVerify::checkExcelFile110(const QString &excelFile1)
{
    memset(&m_excelColSum, 0, sizeof(ExcelFileColSumStruct));
    if(false == readExcelFile1(excelFile1))
    {
        return false;
    }

    int rowCount = m_excelFileRowStructList1.size();

    for(int i = 0; i < rowCount; i++)
    {
        m_excelColSum.gongJiJinYuanGongJiaoNa += m_excelFileRowStructList1[i].geRenGongJiJin;
        m_excelColSum.shiYeBaoXianGongSiJiaoNa += m_excelFileRowStructList1[i].danWeiShiYeJin;
        m_excelColSum.yiLiaoBaoXianGongSiJiaoNa += m_excelFileRowStructList1[i].danWeiYiBao;
        m_excelColSum.shengYuBaoXianGongSiJiaoNa += m_excelFileRowStructList1[i].danWeiShengYuBaoXian;
        m_excelColSum.shiYeBaoXianYuanGongJiaoNa += m_excelFileRowStructList1[i].geRenShiYeJin;
        m_excelColSum.yangLaoBaoXianGongSiJiaoNa += m_excelFileRowStructList1[i].danWeiYangLaoJin;
        m_excelColSum.daBingBaoXianYuanGongJiaoNa += m_excelFileRowStructList1[i].geRenDaBingYiLiao;
        m_excelColSum.yiLiaoBaoXianYuanGongJiaoNa += m_excelFileRowStructList1[i].geRenYiBao;
        m_excelColSum.gongJiJinBaoXianGongSiJiaoNa += m_excelFileRowStructList1[i].danWeiGongJiJin;
        m_excelColSum.gongShangBaoXianGongSiJiaoNa += m_excelFileRowStructList1[i].danWeiGongShangBaoXian;
        m_excelColSum.yangLaoBaoXianYuanGongJiaoNa += m_excelFileRowStructList1[i].geRenYangLaoJin;
    }

    return true;
}


