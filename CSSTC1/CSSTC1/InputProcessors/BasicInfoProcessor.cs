using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Awords = Aspose.Words;
using CSSTC1.ConstantVariables;
using Aspose.Words;
using System.Windows.Forms;
using CSSTC1.FileProcessors;
using CSSTC1.FileProcessors.writers;
using CSSTC1.CommonUtils;

namespace CSSTC1.InputProcessors {
    class BasicInfoProcessor {
        public Document doc = new Document(FileConstants.root_file);
        //public Microsoft.Office.Interop.Word._Application wordApp = new Microsoft.Office.Interop.Word.Application();

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
                }
                index = index + 1;
            }
            string year = values[values.Length - 2].Substring(0, 4);
            doc_builder.MoveToBookmark("年份");
            doc_builder.Write(year);
            this.fill_test_accordings(test_accordings);
            this.change_file_structure(doc, doc_builder);
            doc.Save(FileConstants.save_root_file);
        }

        public void fill_test_accordings(bool[] test_accordings) {
            string ceshiyiju = "";
            for(int i = 0; i < test_accordings.Length; i++){
                bool flag = test_accordings[i];
                if(flag){
                    ceshiyiju += NamingRules.Csyj_files[i] + '、';
                }
            }
            if(ceshiyiju.Length > 0){
                ceshiyiju = ceshiyiju.Substring(0, ceshiyiju.Length - 1);
                Document doc = new Document(FileConstants.save_root_file);
                DocumentBuilder doc_builder = new DocumentBuilder(doc);
            
                if(doc_builder.MoveToBookmark("测试依据")){
                    doc_builder.Write(ceshiyiju);
                }
                doc.Save(FileConstants.save_root_file);
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
            if(ContentFlags.luojiceshi == 0)
                OperationHelper.delete_section(doc, doc_builder, "逻辑测试", "luojiceshi");
            if(ContentFlags.xitonghuiguiceshi == 0)
                OperationHelper.delete_section(doc, doc_builder, "系统回归测试", "xitonghuiguiceshi");
            return true;
        }
        }
    }
