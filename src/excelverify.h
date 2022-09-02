#ifndef EXCELVERIFY_H
#define EXCELVERIFY_H

#include <QString>
#include <QMap>
#include <QVector>

/*
表1：社保公积金结算单
表2：薪资明细表（从系统中导出的）
表3：薪资明细表2.0（造册）
*/
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


class ExcelVerify
{
public:
    ExcelVerify();

    ~ExcelVerify();

    // 表1中标黄的项与表2的相同项，按人逐个比较
    bool checkExcelFile12(const QString &excelFile1, const QString &excelFile2);

    // 表2中标黄的项与表3的相同项，按人逐个比较
    bool checkExcelFile23(const QString &excelFile2, const QString &excelFile3);

    // 表1中标黄的项分别求出总和，并输出在每一列的列尾；同时与界面上手动输入的各项数据进行对比（界面上的数据来源于财务）
    //bool checkExcelFile110(const QString &excelFile1);

private:
    void initMap1();
    void initMap2();
    void initMap3();

    // 读表1
    bool readExcelFile1(const QString &excelFile);

    // 读表2
    bool readExcelFile2(const QString &excelFile);

    // 读表3
    bool readExcelFile3(const QString &excelFile);


private:

    QMap<QString, int> m_map1;
    QMap<QString, int> m_map2;
    QMap<QString, int> m_map3;

    QVector<ExcelFile1RowStruct> m_excelFileRowStructList1;
    QVector<ExcelFile2RowStruct> m_excelFileRowStructList2;
    //QList<ExcelFile3Row> m_excelFileRowList3;
};

#endif // EXCELVERIFY_H
