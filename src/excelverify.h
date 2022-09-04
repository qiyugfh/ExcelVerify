#ifndef EXCELVERIFY_H
#define EXCELVERIFY_H

#include <QString>
#include <QMap>
#include <QVector>


const float CHECK_FLOAT_PRECISION = 0.001f;


// 表1 社保公积金结算单,每行需要检查的列
struct ExcelFile1RowStruct
{
    QString gongHao;
    QString xingMing;
    float   danWeiYangLaoJin;
    float   danWeiYiBao;
    float   danWeiShiYeJin;
    float   danWeiGongShangBaoXian;
    float   danWeiShengYuBaoXian;
    float   danWeiGongJiJin;
    float   geRenYangLaoJin;
    float   geRenYiBao;
    float   geRenShiYeJin;
    float   geRenDaBingYiLiao;
    float   geRenGongJiJin;
};

// 表2 薪资明细表（从系统中导出的）,每行需要检查的列
struct ExcelFile2RowStruct
{
    QString gongHao;
    QString xingMing;
    float   yeBanJinTie;
    float   canTie;
    float   tongXunBuTie;
    float   dianNaoBuTie;
    float   jiaBanBuTie;
    float   qiTaJiangJinZiJinGuiJiZhuanXiangJiLi;
    float   shuiQianQiTaBuFa;
    float   yongJin;
    float   shuiQianQiTaBuKou;
    float   yangLaoBaoXianGongSiJiaoNa;
    float   yiLiaoBaoXianGongSiJiaoNa;
    float   shengYuBaoXianGongSiJiaoNa;
    float   shiYeBaoXianGongSiJiaoNa;
    float   gongShangBaoXianGongSiJiaoNa;
    float   gongJiJinGongSiJiaoNa;
    float   yangLaoBaoXianYuanGongJiaoNa;
    float   yiLiaoBaoXianYuanGongJiaoNa;
    float   shiYeBaoXianYuanGongJiaoNa;
    float   daBingBaoXianYuanGongJiaoNa;
    float   gongJiJinYuanGongJiaoNa;
};

// 表3 薪资明细表2.0（造册）,每行需要检查的列
struct ExcelFile3RowStruct
{
    QString gongHao;
    QString xingMing;
    float   tongXunBuTie;
    float   dianNaoBuTie;
    float   yongJinJiLi;
    float   yeBanBuTie;
    float   jiaBanBuTie;
    float   qiTaBuTie;
    float   buFaBuKou;
};


struct CheckCellResult
{
    float val;
    bool  error;
};


struct ExcelFileCheckErrorRowStruct
{
    QString gongHao;
    QString xingMing;
    QVector<CheckCellResult> checkCellResults;
};

/*
养老保险公司缴纳
医疗保险公司缴纳
生育保险公司缴纳
失业保险公司缴纳
工伤保险公司缴纳
公积金公司缴纳
养老保险员工缴纳
医疗保险员工缴纳
失业保险员工缴纳
大病保险员工缴纳
公积金员工缴纳
*/
struct ExcelFileColSumStruct
{
    float yangLaoBaoXianGongSiJiaoNa;
    float yiLiaoBaoXianGongSiJiaoNa;
    float shengYuBaoXianGongSiJiaoNa;
    float shiYeBaoXianGongSiJiaoNa;
    float gongShangBaoXianGongSiJiaoNa;
    float gongJiJinBaoXianGongSiJiaoNa;
    float yangLaoBaoXianYuanGongJiaoNa;
    float yiLiaoBaoXianYuanGongJiaoNa;
    float shiYeBaoXianYuanGongJiaoNa;
    float daBingBaoXianYuanGongJiaoNa;
    float gongJiJinYuanGongJiaoNa;
};


class ExcelVerify
{
public:
    ExcelVerify();

    ~ExcelVerify();

    // 表1中标黄的项与表2的相同项，按人逐个比较
    bool checkExcelFile12(const QString &excelFile1, const QString &excelFile2);

    // 表2中标黄的项与表3的相同项，按人逐个比较
    bool checkExcelFile23(const QString &excelFile2, const QString &excelFile3);

    // 表1中标黄的项分别求出总和
    bool checkExcelFile110(const QString &excelFile1);

public:
    QVector<ExcelFileCheckErrorRowStruct> m_checkResultList1;
    QVector<ExcelFileCheckErrorRowStruct> m_checkResultList2;
    ExcelFileColSumStruct m_excelColSum;

private:

    // 读表1
    bool readExcelFile1(const QString &excelFile);

    // 读表2
    bool readExcelFile2(const QString &excelFile);

    // 读表3
    bool readExcelFile3(const QString &excelFile);

    bool checkFloatValueIsDiff(float val1, float val2);

    bool writeCheckResult1ToFile();


private:

    QVector<ExcelFile1RowStruct> m_excelFileRowStructList1;
    QVector<ExcelFile2RowStruct> m_excelFileRowStructList2;
    QVector<ExcelFile3RowStruct> m_excelFileRowStructList3;

};

#endif // EXCELVERIFY_H
