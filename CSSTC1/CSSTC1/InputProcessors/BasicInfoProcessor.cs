using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Awords = Aspose.Words;
using CSSTC1.ConstantVariables;
using Aspose.Words;
using System.Windows.Forms;

namespace CSSTC1.InputProcessors {
    class BasicInfoProcessor {
        public Document doc = new Document(FilePaths.root_file);
        public Microsoft.Office.Interop.Word._Application wordApp = new Microsoft.Office.Interop.Word.Application();

        public void fill_basic_info(string[] bookmarks, string[] values){
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
            doc.Range.UpdateFields();
            doc.Save(FilePaths.save_cache_file, SaveFormat.Docx);
            this.update_document(FilePaths.save_cache_file);
        }

        public void update_document(string path) {
            this.wordApp.Visible = false;
            this.wordApp.Documents.Open(path);
            object oTemplate = path;
            Microsoft.Office.Interop.Word._Document oDoc = this.wordApp.Documents.Add(ref oTemplate, ref                                                                    FilePaths.oMissing, ref FilePaths.oMissing, ref FilePaths.oMissing);
            oDoc.Fields.Update();
            oDoc.SaveAs2(FilePaths.save_root_file);
            this.wordApp.Quit();
        }
        }
    }
