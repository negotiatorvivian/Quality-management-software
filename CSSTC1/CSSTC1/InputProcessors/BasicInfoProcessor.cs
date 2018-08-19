using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSSTC1.ConstantVariables;
using Aspose.Words;
using System.Windows.Forms;
using CSSTC1.FileProcessors;
using CSSTC1.FileProcessors.writers;
using CSSTC1.CommonUtils;

namespace CSSTC1.InputProcessors {
    class BasicInfoProcessor {
        public Document doc = new Document(FileConstants.root_file);

        public void fill_basic_info(string[] bookmarks, string[] values, bool[] test_accordings) {
            DocumentBuilder doc_builder = new DocumentBuilder(this.doc);
            int index = 0;
            BindingFlags flag = BindingFlags.Static | BindingFlags.Public;
            foreach(string bookmark in bookmarks) {
                FieldInfo f_key = typeof(NamingRules).GetField(bookmark, flag);
                if(f_key != null){
                    string o = f_key.GetValue(new NamingRules()).ToString();

                    if(doc_builder.MoveToBookmark(o))
                        doc_builder.Write(values[index]);
                    if(o.Equals("项目标识")){
                        if(doc_builder.MoveToBookmark("项目标识正文"))
                            doc_builder.Write(values[index]);
                    }
                }
                index = index + 1;
            }
            string year = values[values.Length - 2].Substring(0, 4);
            doc_builder.MoveToBookmark("年份");
            doc_builder.Write(year);
            doc_builder.MoveToBookmark("年份正文");
            doc_builder.Write(year);
            this.fill_test_accordings(doc, doc_builder, test_accordings);
            this.change_file_structure(doc, doc_builder);
            doc.Save(FileConstants.save_root_file);
        }

        public void fill_test_accordings(Document doc, DocumentBuilder doc_builder, bool[] test_accordings) {
            string ceshiyiju = "";
            for(int i = 0; i < test_accordings.Length; i++){
                bool flag = test_accordings[i];
                if(flag){
                    ceshiyiju += NamingRules.Csyj_files[i] + '、';
                }
            }
            if(ceshiyiju.Length > 0){
                ceshiyiju = ceshiyiju.Substring(0, ceshiyiju.Length - 1);
                if(doc_builder.MoveToBookmark("测试依据")){
                    doc_builder.Write(ceshiyiju);
                }
            }
       }

        public bool change_file_structure(Document doc, DocumentBuilder doc_builder){
            if(ContentFlags.wendangshencha == 0)
                OperationHelper.delete_section(doc, doc_builder, "文档审查", "wendangshencha");
            if(ContentFlags.jingtaifenxi == 0)
                OperationHelper.delete_section(doc, doc_builder, "静态分析", "jingtaifenxi");
            if(ContentFlags.daimashencha == 0)
                OperationHelper.delete_section(doc, doc_builder, "代码审查", "daimashencha");
            if(ContentFlags.daimazoucha == 0)
                OperationHelper.delete_section(doc, doc_builder, "代码走查", "daimazoucha");
            if(ContentFlags.peizhiceshi == 0)
                OperationHelper.delete_section(doc, doc_builder, "配置项测试", "peizhiceshi");
            //if(ContentFlags.luojiceshi == 0)
            //    OperationHelper.delete_section(doc, doc_builder, "逻辑测试", "luojiceshi");
            if(ContentFlags.xitongceshi == 0)
                OperationHelper.delete_section(doc, doc_builder, "系统测试", "xitongceshi");
            else if(ContentFlags.xitonghuiguiceshi == 0)
                OperationHelper.delete_section(doc, doc_builder, "系统回归测试", "xitonghuiguiceshi");
            return true;
        }
        }
    }
