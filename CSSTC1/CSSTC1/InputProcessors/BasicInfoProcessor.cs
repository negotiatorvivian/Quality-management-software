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

namespace CSSTC1.InputProcessors {
    class BasicInfoProcessor {
        public Document doc = new Document(FilePaths.root_file);
        //public Microsoft.Office.Interop.Word._Application wordApp = new Microsoft.Office.Interop.Word.Application();
        public FileWriters1 file_writer = new FileWriters1();

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
            NodeCollection nodes = doc.GetChildNodes(NodeType.FieldStart, true);
            foreach(Aspose.Words.Fields.FieldStart field_ref in nodes){
                Aspose.Words.Fields.Field field = field_ref.GetField();
                field.Update();
                //doc.Save(FilePaths.save_root_file);
            }
            doc.Save(FilePaths.save_root_file);
            fill_test_accordings(test_accordings);
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
                Document doc = new Document(FilePaths.save_root_file);
                DocumentBuilder doc_builder = new DocumentBuilder(doc);
            
                if(doc_builder.MoveToBookmark("测试依据")){
                    doc_builder.Write(ceshiyiju);
                }
                doc.Save(FilePaths.save_root_file);
            }
       }
        }
    }
