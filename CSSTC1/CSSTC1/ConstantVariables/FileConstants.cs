﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.ConstantVariables {
    class FileConstants {
        public static string root_file = @"C:\Users\vivian\Documents\自定义 Office 模板\CSSTC模板.dotx";
        public static string save_cache_file = @"D:\张紫薇\实习\CSSTC\cache\CSSTC.docx";
        public static string save_root_file = @"D:\张紫薇\实习\CSSTC\CSSTC.docx";

        public static object oMissing = System.Reflection.Missing.Value;

        //文档审查节数
        public static int wendangshencha = 6;
        //静态分析节数
        public static int jingtaifenxi = 15;
        //代码审查节数
        public static int daimashencha = 13;
        //是否有第二次测试就绪
        public static int ceshijiuxu2 = 5;
        //代码走查节数
        public static int daimazoucha = 13;
        //配置项测试节数
        public static int peizhiceshi = 23;
        //逻辑测试节数
        public static int luojiceshi = 3;
        //系统测试节数
        public static int xitongceshi = 20;
        //系统回归测试节数
        public static int xitonghuiguiceshi = 10;
    }
}