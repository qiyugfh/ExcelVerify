# 本工具软件的需求说明：
表1：社保公积金结算单
表2：薪资明细表（从系统中导出的）
表3：薪资明细表2.0（造册）

需求一：
表1中标黄的项与表2的相同项，按人逐个比较；如果有不同，则给出提示，并把不同的项高亮显示；

需求二：
表2中标黄的项与表3的相同项，按人逐个比较；如果有不同，则给出提示，并把不同的项高亮显示；

需求三：
表1中标黄的项分别求出总和，并输出在每一列的列尾；同时与界面上手动输入的各项数据进行对比（界面上的数据来源于财务）。如果不同，则给出提示，并把不同的项高亮显示。


薪资明细表 AI（夜班津贴）=薪资明细表2.0 O（夜班补贴）
薪资明细表 AS（通讯补贴）=薪资明细表2.0 L（通讯补贴）
薪资明细表 AT（电脑补贴）=薪资明细表2.0 M（电脑补贴）
薪资明细表 AU（加班补贴）=薪资明细表2.0 P（加班补贴）
薪资明细表 BQ（其他奖金/资金归集专项激励）+薪资明细表 AR（餐贴）=薪资明细表2.0 Q（其他补贴）
薪资明细表 BW（佣金）=薪资明细表2.0 N（佣金激励）
薪资明细表 BS（税前其他补发）+薪资明细表 CB（税前其他补扣）=薪资明细表2.0 R（补发补扣）
