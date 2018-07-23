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
        public void fill_basic_info(string[] bookmarks, string[] values){
            DocumentBuilder doc_builder = new DocumentBuilder(this.doc);
            int index = 0;
            BindingFlags flag = BindingFlags.Static | BindingFlags.Public;
            foreach(string bookmark in bookmarks) {
                FieldInfo f_key = typeof(NamingRules).GetField(bookmark, flag);
                if(f_key != null){
                    string o = f_key.GetValue(new NamingRules()).ToString();
                    //MessageBox.Show(o);
                    doc_builder.MoveToBookmark(o);
                    doc_builder.Write(values[index]);
                }
                
                index = index + 1;
            }
            doc.Range.UpdateFields();
            doc.Save(FilePaths.save_root_file, SaveFormat.Docx);

        }

        
        }
    }
