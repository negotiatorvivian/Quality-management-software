using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CSSTC1.CommonUtils {
    class MappingHelper {
        public static string get_file_type(string name) {
            Regex reg = new Regex(@"[\u4e00-\u9fa5]{4,}");
            Match match = reg.Match(name, 0);
            string title = match.Groups[0].Value;
            switch(title) {
                case "软件研制总要求":
                    return "01";
                case "软件研制任务书":
                    return "02";
                case "软件任务书":
                    return "02";
                case "任务书":
                    return "02";
                case "软件质量保证计划":
                    return "03";
                case "软件配置管理计划":
                    return "04";
                case "软件需求规格说明":
                    return "05";
                case "需求规格说明":
                    return "05";
                case "接口需求规格说明":
                    return "06";
                case "软件设计文档":
                    return "07";
                case "设计文档":
                    return "07";
                case "软件概要设计文档":
                    return "08";
                case "软件概要设计报告":
                    return "08";
                case "软件详细设计文档":
                    return "09";
                case "软件详细设计说明书":
                    return "09";
                case "详细设计说明书":
                    return "09";
                case "详细设计文档":
                    return "09";
                case "接口设计文档":
                    return "10";
                case "接口设计":
                    return "10";
                case "软件产品规格说明":
                    return "11";
                case "版本说明文档":
                    return "12";
                case "软件测试计划":
                    return "13";
                case "软件测试说明":
                    return "14";
                case "软件测试报告":
                    return "15";
                case "计算机系统操作员手册":
                    return "16";
                case "软件用户手册":
                    return "17";
                case "软件程序员手册":
                    return "18";
                case "软件代码":
                    return "19";
                case "被测系统硬件":
                    return "20";
                case "测试协议":
                    return "21";
                case "测试合同":
                    return "22";
                case "被测件清单":
                    return "23";
                case "被测件光盘":
                    return "24";
                case "软件安装包":
                    return "26";
                default:
                    return "25";
            }

        }

        public static string get_job_title(string name) {
            switch(name) {
                case "陈大圣":
                    return "主任";
                case "唐龙利":
                    return "常务副主任";
                case "韩新宇":
                    return "副主任";
                case "张凯":
                    return "高工";
                case "贺仁亚":
                    return "副研究员";
                default:
                    return "工程师";
            }
        }

        public static int get_doc_section(string name) {
            switch(name) {
                case "被测件清单":
                    return 3;
                case "配置报告单":
                    return 4;
                default:
                    return 0;
            }
        }

    }
}
