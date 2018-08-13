using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.ConstantVariables {
    class InsertionPos {
        //全局变量
        public static int hyqdb_name_row = 1;
        public static int hyqdb_company_row = 2;
        public static int hyqdb_job_row = 3;

        public static int pz_table = 2;
        public static int pz_row = 3;
        public static int pz_cell0 = 0;
        public static int pz_cell1 = 0;

        public static int xt_table = 2;
        public static int xt_row = 4;
        public static int xt_cell0 = 0;
        public static int xt_cell1 = 1;

        public static int csry_table = 3;
        public static int csry_row = 5;
        public static int csry_cell = 2;

        public static int jbqk_table = 1;
        public static int jbqk_bcyy_row = 4;
        public static int jbqk_bchj_row = 5;
        public static int jbqk_xtqk_row = 8;
        public static int jbqk_cell = 1;

        //测试合同内审签到表
        public static int cshtqd_section = 0;
        public static int cshtqd_sec_table = 5;

        public static int bcjdbd_table = 7;
        public static int bcjdbd_res_row = 1;
        public static int bcjdbd_name_row = 2;
        public static int bcjdbd_pois_row = 6;
        public static int bcjdbd_sec_table = 0;

        public static int lx_bcjqd_table = 8;
        public static int lx_bcjqd_sec_table = 1;
        public static int lx_bcjqd_res_row = 3;
        public static int lx_bcjqd_name_row = 1;
        public static int lx_bcjqd_orig_row = 7;
        public static int lx_bcjqd_date_row = 8;

        public static int bcjlqqd_table = 9;
        public static int bcjlqqd_sec_table = 2;
        public static int bcjlqqd_res_row = 1;
        public static int bcjlqqd_name_row = 2;

        

        

        public static int rksqd_table = 10;
        public static int rksqd_sec_table = 3;
        public static int rksqd_name_row = 0;
        public static int rksqd_iden_row = 3;
        
        //测试大纲内审表位置
        public static int csdgns_sec_table1 = 5;
        public static int csdgns_ques_row = 1;
        public static int csdgns_solu_row = 3;
        public static int csdgns_sec_table2 = 6;
        public static int csdgns_section = 2;
        //会议签到表
        public static int csdgqd_section = 2;
        public static int csdgqd_sec_table = 4;


        //测试大纲外审表位置
        public static int csdgws_sec_table1 = 2;
        public static int csdgws_sec_table2 = 3;
        public static int csdgws_section = 4;
        public static int csdgws_ques_row = 1;
        public static int csdgws_solu_row = 3;

        //测试需求内审表位置
        public static int csxqns_sec_table1 = 2;
        public static int csxqns_sec_table2 = 3;
        public static int csxqns_section = 2;
        public static int csxqns_ques_row = 1;
        public static int csxqns_solu_row = 3;
        //会议签到表
        public static int csxqqd_section = 2;
        public static int csxqqd_sec_table = 1;

        //测试策划内审表位置
        public static int cschns_sec_table1 = 6;
        public static int cschns_sec_table2 = 7;
        public static int cschns_section = 2;
        public static int cschns_ques_row = 1;
        public static int cschns_solu_row = 3;
        //会议签到表
        public static int cschqd_section = 2;
        public static int cschqd_sec_table = 5;

        //测试需求与策划外审表位置
        public static int csxqws_sec_table1 = 2;
        public static int csxqws_sec_table2 = 3;
        public static int csxqws_section = 4;
        public static int csxqws_ques_row = 1;
        public static int csxqws_solu_row = 3;

        //配置管理计划表位置
        public static int pzgljh_section = 2;
        public static int pzgljh_sec_table = 1;
        public static int pzgljh_name_row = 0;
        public static int pzgljh_iden_row = 1;
        /******************************文档审查部分**************************/
        //测试设计与执行第一个被测件清单位置
        public static int sj_bcjqd_section = 11;
        public static int sj_bcjqd_sec_table = 0;
        public static int sj_bcjqd_name_row = 1;
        public static int sj_bcjqd_iden_row = 3;
        public static int sj_bcjqd_orig_row = 7;
        public static int sj_bcjqd_date_row = 8;
        //测试设计与执行第一个被测件调拨单位置
        public static int sj_bcjdbd_section = 11;
        public static int sj_bcjdbd_sec_table = 1;
        public static int sj_bcjdbd_name_row = 2;
        public static int sj_bcjdbd_iden_row = 1;
        //测试设计与执行第一个被测件领取清单位置
        public static int sj_bcjlqqd_sec_table = 2;

        //测试设计与执行第一个入库申请单位置
        public static int sj_rksqd_section = 13;
        public static int sj_rksqd_sec_table = 0;
        public static int sj_rksqd_name_row = 0;
        public static int sj_rksqd_iden_row = 3;
        /******************************文档审查部分结束**************************/

        /******************************静态分析部分开始**************************/
        //测试设计与执行第二个被测件调拨单位置
        public static int sj_bcjdbd_section2 = 10;
        public static int sj_bcjdbd_sec_table2 = 0;

        //测试设计与执行第二个被测件清单位置
        public static int sj_bcjqd_section2 = 10;
        public static int sj_bcjqd_sec_table2 = 1;
        //测试设计与执行第二个被测件领取清单位置
        public static int sj_bcjlqqd_section2 = 10;
        public static int sj_bcjlqqd_sec_table2 = 2;
        //测试设计与执行第二个入库申请单位置
        public static int sj_rksqd_section2 = 10;
        public static int sj_rksqd_sec_table2 = 3;
        //编程语言列
        public static int sj_code_row_index = 2;

        //测试设计与执行第一个测试工具或设备确认表位置
        public static int sj_csgjqr_section = 12;
        public static int sj_csgjqr_sec_table = 0;
        public static int sj_csgjqr_name_row = 2;
        public static int sj_csgjqr_iden_row = 4;

        //测试设计与执行测试环境核查表位置
        public static int sj_cshjhc_section = 13;
        public static int sj_cshjhc_sec_table = 0;
        //测试设计与执行第三个被测件清单位置
        public static int sj_bcjqd_section3 = 18;
        public static int sj_bcjqd_sec_table3 = 0;
        //测试设计与执行第三个被测件调拨单位置
        public static int sj_bcjdbd_section3 = 18;
        public static int sj_bcjdbd_sec_table3 = 1;
        //测试设计与执行第三个被测件领取清单位置
        public static int sj_bcjlqqd_section3 = 18;
        public static int sj_bcjlqqd_sec_table3 = 2;
        //测试设计与执行第三个入库申请单位置
        public static int sj_rksqd_section3 = 18;
        public static int sj_rksqd_sec_table3 = 3;
        /******************************静态分析部分结束**************************/

        /******************************代码审查/走查部分开始**************************/
        public static int dmsc_bcjqd_section = 10;
        public static int dmsc_bcjqd_sec_table = 0;
        public static int dmsc_bcjdbd_sec_table1 = 1;
        public static int dmsc_bcjlqqd_sec_table1 = 2;
        public static int dmsc_rksqd_sec_table1 = 3;

        public static int dmsc_chhjhcd_section = 12;
        public static int dmsc_cshj_sec_table1 = 0;
        public static int dmsc_rjhj_sec_table1 = 1;
        public static int dmsc_rjhj_name_row = 1;
        public static int dmsc_rjhj_version_row = 2;
        public static int dmsc_rjhj_orig_row = 4;

        public static int dmsc_bcjqd_section1 = 16;
        /******************************代码审查/走查部分结束**************************/ 
   
        /******************************测试说明评审部分开始**************************/
        //测试说明内审表偏差问题报告
        public static int cssmns_section = 9;
        public static int cssmns_sec_table1 = 2;
        //会议签到表
        public static int cssmns_hyqdb_table = 1;
        //测试说明内审表偏差与问题追踪
        public static int cssmns_sec_table2 = 3;
        public static int cssmns_ques_row = 1;
        public static int cssmns_solu_row = 3;

        //测试说明评审入库申请单位置
        public static int cssmps_rksqd_section = 10;
        public static int cssmps_rksqd_sec_table = 0;
        //测试说明评审被测件清单位置
        public static int cssmps_bcjqd_section = 13;
        public static int cssmps_bcjqd_sec_table = 0;
        public static int cssmns_bcjqd_name_row = 1;
        public static int cssmns_bcjqd_iden_row = 3;
        //测试说明评审被测件清单位置
        public static int cssmps_bcjdbd_sec_table = 1;
        //测试说明评审被测件领取清单位置
        public static int cssmps_bcjlqqd_sec_table = 2;
        /******************************测试说明评审部分结束**************************/

        /******************************搭建环境部分开始**************************/
        //搭建环境阶段被测件清单第一张表格位置
        public static int djhj_bcjqd_section = 13;
        public static int djhj_bcjqd_sec_table = 0;
        public static int djhj_bcjdbd_sec_table = 1;
        public static int djhj_bcjlqqd_sec_table = 2;
        public static int djhj_bcjlqqd_type_row = 6;

        //搭建环境阶段测试环境确认表位置
        public static int djhj_cshjqr_section = 15;
        public static int djhj_cshjqr_sec_table = 0;
        //搭建环境阶段测试环境核查表位置
        public static int djhj_cshjhc_section = 16;
        public static int djhj_cshjhc_sec_table = 1;
        //测试就绪评审会议签到表
        public static int djhj_hyqdb_section = 15;
        public static int djhj_hyqdb_sec_table = 2;
        public static int djhj_pcywt_sec_table = 0;
        public static int djhj_pcywtzz_sec_table = 1;

        //搭建环境阶段被测件清单第二张表格位置
        public static int djhj_bcjqd_section1 = 16;
        //搭建环境阶段第二次测试环境确认表位置
        public static int djhj_cshjqr_section1 = 18;
        //搭建环境阶段第二次测试环境核查表位置
        public static int djhj_cshjhc_section1 = 19;
        //测试就绪评审会议签到表2
        public static int djhj_hyqdb_section1 = 18;
        /******************************搭建环境部分结束**************************/

        /******************************配置测试部分开始**************************/
        //首轮配置项测试第一张入库申请表格位置
        public static int pzcs_rksqd_section = 17;
        public static int pzcs_rksqd_sec_table = 0;
        //配置项回归测试第一张被测件清单位置
        public static int pzcs_bcjqd_section = 19;
        public static int pzcs_bcjqd_sec_table = 0;
        public static int pzcs_bcjdbd_sec_table = 1;
        public static int pzcs_bcjlqqd_sec_table = 2;

        //配置项回归测试第一次测试环境确认表位置
        public static int pzcs_cshjqr_section1 = 21;
        public static int pzcs_cshjhc_section1 = 22;
        //配置项回归测试第二张入库申请表位置
        public static int pzcs_rksqd_section1 = 22;
        public static int pzcs_rksqd_section2 = 26;
        /******************************配置测试部分结束**************************/

        /******************************系统测试部分开始**************************/
        public static int xtcs_section = 17;
        public static int xtcs_bcjqd_sec_table = 0;
        public static int xtcs_bcjdbd_sec_table = 1;
        //系统回归测试环境确认单
        public static int xtcs_cshjqr_section1 = 19;
        public static int xtcs_cshjhc_section1 = 20;
        public static int xtcs_csgjqr_sec_table = 0;
        public static int xtcs_csgjhc_sec_table = 0;
        //系统回归测试被测件清单
        public static int xtcs_section1 = 27;
        //系统回归测试环境确认单
        public static int xtcs_cshjqr_section2 = 29;
        public static int xtcs_cshjhc_section2 = 30;
        /******************************系统测试部分结束**************************/

        /******************************项目总结部分开始**************************/
        //内审意见表
        public static int xtzj_jdcp_section = 17;
        public static int xtzj_hyqdb_sec_table = 1;
        public static int xtzj_wtbg_sec_table = 2;
        public static int xtzj_wtzz_sec_table = 3;
        //外审意见表
        public static int xtzj_jdcp_section1 = 19;
        //出库申请单
        public static int xtzj_cksqd_section = 24;
        public static int xtzj_cksqd_section1 = 26;
        public static int xtzj_cksqd_sec_table = 0;
        //第一个出库申请单文档审查记录起始位置
        public static int xtzj_wdscjl_row_index1 = 8;
        //第二个出库申请单文档审查记录起始位置
        public static int xtzj_wdscjl_row_index2 = 6;

        //第一个入库申请单软件测试执行记录附件起始位置
        public static int xtzj_cszxjl_row_index1 = 9;
        //第二个入库申请单软件测试执行记录附件起始位置
        public static int xtzj_cszxjl_row_index2 = 7;

        //第二个出库申请单测试说明附件起始位置
        public static int xtzj_cssmfj_row_index2 = 5;
        
        /******************************项目总结部分结束**************************/
    }
}
