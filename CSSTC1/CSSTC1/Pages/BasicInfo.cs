using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CSSTC1.InputProcessors;
using CSSTC1.ConstantVariables;
using CSSTC1.CommonUtils;
using Aspose.Words;

namespace CSSTC1.Pages {
    public partial class BasicInfo : Form {
        public BasicInfo() {
            InitializeComponent();
            }
        #region  定义变量及日期点击事件
        public string Xm_biaoshi;
        public string Xm_mingcheng;
        public string Rj_mingcheng;
        public bool Pzx_ceshi;
        public bool Xt_ceshi;
        public string Pz_guanliyuan;
        public string Bcj_guanliyuan;
        public string Zlbz_renyuan;
        public string Cs_jianduyuan;
        public string Sb_guanliyuan;
        public string Fuzhurren;
        public string Cw_fuzhuren;
        public string Zhuren;
        public string Wt_xingming;
        public string Wt_dianhua;
        public string Wt_danwei;
        public string Kf_xingming;
        public string Kf_dianhua;
        public string Kf_danwei;
        public string Xm_kaishishijian;
        public string Xm_jieshushijian;

        //测试依据文档
        public bool Yz_renwushu;
        public bool Js_fangan;
        public bool Js_guigeshu;
        public bool Rjyz_renwushu;
        public bool Rjxq_guigeshuoming;

        BasicInfoProcessor processor = new BasicInfoProcessor();

        private void checkBox8_CheckedChanged(object sender, EventArgs e) {
            this.xtcs_huiguishijian.Enabled = this.checkBox8.Checked;
        }

        //配置项首轮动态测试时间
        private void dtpz_shoulunshijian_ValueChanged(object sender, EventArgs e) {
            TimeStamp.sldtcs_format_time = this.dtpz_shoulunshijian.Value.ToLongDateString();
            TimeStamp.sldtcs_time = this.dtpz_shoulunshijian.Value.ToShortDateString();
        }

        //系统首轮测试时间
        private void xtcs_shoulunshijian_ValueChanged(object sender, EventArgs e) {
            TimeStamp.slxtcs_time = this.xtcs_shoulunshijian.Value.ToShortDateString();
            TimeStamp.slxtcs_format_time = this.xtcs_shoulunshijian.Value.ToLongDateString();
        }

        //测试报告评审时间
        private void csbg_pingshenshijian_ValueChanged(object sender, EventArgs e) {
            TimeStamp.csbgps_time = this.csbg_pingshenshijian.Value.ToShortDateString();
            TimeStamp.csbgps_format_time = this.csbg_pingshenshijian.Value.ToLongDateString();
        }
        #endregion

        #region  界面组件状态控制
        private void button2_Click(object sender, EventArgs e) {
            Globals.ThisDocument.basic_info.Hide();
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e) {
            bool status = this.checkBox6.Enabled;
            this.checkBox6.Checked = !status;
            this.checkBox6.Enabled = !status;
            this.csdg_pingshenshijian.Enabled = !status;
            this.checkBox7.Checked = status;
            this.checkBox7.Enabled = status;
            this.csxq_pingshenshijian.Enabled = status;

        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e) {
            bool status = this.checkBox7.Enabled;
            this.checkBox6.Checked = status;
            this.checkBox6.Enabled = status;
            this.csdg_pingshenshijian.Enabled = status;
            this.checkBox7.Checked = !status;
            this.checkBox7.Enabled = !status;
            this.csxq_pingshenshijian.Enabled = !status;
        }

        private void xm_mingcheng_TextChanged(object sender, EventArgs e) {
            string temp = this.xm_mingcheng.Text;
            if(temp.Length > 0) {
                int index = temp.IndexOf("软件");
                if(index > 0)
                    temp = temp.Substring(0, index + 2);
                this.rj_mingcheng.Text = temp;
            }
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e) {
            foreach(Control control in this.panel2.Controls)
                control.Enabled = this.checkBox9.Checked;
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e) {
            this.panel3.Enabled = this.checkBox10.Checked;
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e) {
            this.panel4.Enabled = this.checkBox11.Checked;
        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e) {
            this.panel5.Enabled = this.checkBox12.Checked;
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e) {
            this.ljcs_shijian.Enabled = this.checkBox4.Checked;
        }
        #endregion

        //提交信息
        private void button1_Click(object sender, EventArgs e) {
            button1.Enabled = false;
            this.Xm_biaoshi = xm_biaoshi.Text;
            this.Xm_mingcheng = xm_mingcheng.Text;
            this.Rj_mingcheng = rj_mingcheng.Text;
            if(!pzx_ceshi.Checked)
                ContentFlags.peizhiceshi = 0;
            if(!xt_ceshi.Checked)
                ContentFlags.xitongceshi = 0;
            this.Pz_guanliyuan = pz_guanliyuan.Text;
            this.Bcj_guanliyuan = bcj_guanliyuan.Text;
            this.Zlbz_renyuan = zlbz_renyuan.Text;
            this.Cs_jianduyuan = cs_jianduyuan.Text;
            this.Sb_guanliyuan = sb_guanliyuan.Text;
            this.Fuzhurren = fuzhurren.Text;
            this.Cw_fuzhuren = cw_fuzhuren.Text;
            this.Zhuren = zhuren.Text;
            this.Wt_xingming = wt_xingming.Text;
            this.Wt_dianhua = wt_dianhua.Text;
            this.Wt_danwei = wt_danwei.Text;
            this.Kf_xingming = kf_xingming.Text;
            this.Kf_dianhua = kf_dianhua.Text;
            this.Kf_danwei = kf_danwei.Text;
            this.Xm_kaishishijian = xm_kaishishijian.Value.ToLongDateString();
            TimeStamp.xiangmukaishishijian = xm_kaishishijian.Value.ToShortDateString();
            this.Xm_jieshushijian = xm_jieshushijian.Value.ToLongDateString();                                                 ContentFlags.ceshidagang = this.checkBox6.Checked;
            this.Yz_renwushu = checkBox1.Checked;
            this.Js_fangan = checkBox2.Checked;
            this.Js_guigeshu = checkBox3.Checked;
            this.Rjyz_renwushu = checkBox4.Checked;
            this.Rjxq_guigeshuoming = checkBox5.Checked;

            

            string[] values = { Xm_biaoshi, Xm_mingcheng, Rj_mingcheng, Pz_guanliyuan, Bcj_guanliyuan, 
                                  Zlbz_renyuan, Cs_jianduyuan, Sb_guanliyuan, Fuzhurren, Cw_fuzhuren, Zhuren,
                                  Wt_xingming, Wt_dianhua, Wt_danwei, Kf_xingming, Kf_dianhua, Kf_danwei, 
                                  Xm_kaishishijian, Xm_jieshushijian};
            string[] bookmarks = { "Xm_biaoshi", "Xm_mingcheng", "Rj_mingcheng", "Pz_guanliyuan", "Bcj_guanliyuan",
                                     "Zlbz_renyuan","Cs_jianduyuan", "Sb_guanliyuan", "Fuzhurren", "Cw_fuzhuren", 
                                     "Zhuren", "Wt_xingming","Wt_dianhua", "Wt_danwei", "Kf_xingming", 
                                     "Kf_dianhua", "Kf_danwei", "Xm_kaishishijian", "Xm_jieshushijian" };
            bool[] test_accordings = { this.Yz_renwushu, this.Js_fangan, this.Js_guigeshu, this.Rjyz_renwushu, 
                                         this.Rjxq_guigeshuoming };
            this.record_time();
            processor.fill_basic_info(bookmarks, values, test_accordings);
            this.del_section();
            Globals.ThisDocument.basic_info.Hide();
            //MessageBox.Show("正在写入，请稍候...");
        }

        

        //大纲与需求说明二选一，删除未选的表格部分
        private void del_section(){
            if(!this.csdg_pingshenshijian.Enabled){
                string[] bookmarks = { "测试大纲1", "测试大纲2", "测试大纲3", "测试大纲4", 
                                         "测试大纲5", "测试大纲6", "测试大纲7"};
                CommonUtils.OperationHelper.delete_section(FileConstants.save_root_file, bookmarks);
            }
            else{
                string[] bookmarks = { "测试需求1", "测试需求2", "测试需求3", "测试需求4", 
                                         "测试需求5", "测试需求6", "测试需求7"};
                CommonUtils.OperationHelper.delete_section(FileConstants.save_root_file, bookmarks);
            }
        }

        //所有时间信息
        public void record_time(){
            if(this.checkBox6.Checked) {
                //测试大纲评审时间
                TimeStamp.ceshisc_time = csdg_pingshenshijian.Value.ToShortDateString();
                //测试大纲评审时间文字形式
                TimeStamp.ceshisc_format_time = csdg_pingshenshijian.Value.ToLongDateString();
            }
            if(this.checkBox7.Checked) {
                //测试大纲或测试说明评审时间
                TimeStamp.ceshisc_time = csxq_pingshenshijian.Value.ToShortDateString();
                //测试大纲或测试说明评审时间文字形式
                TimeStamp.ceshisc_format_time = csxq_pingshenshijian.Value.ToLongDateString();
            }
            if(this.checkBox9.Checked){
                //文档审查时间            
                TimeStamp.wdscqr_time = this.wd_querenshijian.Value.ToShortDateString();
                TimeStamp.wdscqr_format_time = this.wd_querenshijian.Value.ToLongDateString();
                TimeStamp.wdschg_time = this.wd_huiguishijian.Value.ToShortDateString();
                TimeStamp.wdschg_format_time = this.wd_huiguishijian.Value.ToLongDateString();
            }
            else
                ContentFlags.wendangshencha = 0;
            if(this.checkBox10.Checked){
            //静态分析时间
                TimeStamp.jtfxsc_time = this.dateTimePicker2.Value.ToShortDateString();
                TimeStamp.jtfxsc_format_time = this.dateTimePicker2.Value.ToLongDateString();
                TimeStamp.jtfxhg_time = this.jt_huiguishijian.Value.ToShortDateString();
                TimeStamp.jtfxhg_format_time = this.jt_huiguishijian.Value.ToLongDateString();
                TimeStamp.jtfxqr_time = this.dateTimePicker1.Value.ToShortDateString();
                TimeStamp.jtfxqr_format_time = this.dateTimePicker1.Value.ToLongDateString();
            }
            else
                ContentFlags.jingtaifenxi = 0;
            if(this.checkBox11.Checked) {
            //代码审查时间
                TimeStamp.dmsc_time = this.dmsc_shijian.Value.ToShortDateString();
                TimeStamp.dmsc_format_time = this.dmsc_shijian.Value.ToLongDateString();
                TimeStamp.dmscqr_time = this.dateTimePicker3.Value.ToShortDateString();
                TimeStamp.dmscqr_format_time = this.dateTimePicker3.Value.ToLongDateString();
                TimeStamp.dmschg_time = this.dmhg_huiguishijian.Value.ToShortDateString();
                TimeStamp.dmschg_format_time = this.dmhg_huiguishijian.Value.ToLongDateString();
            }
            else
                ContentFlags.daimashencha = 0;
            if(this.checkBox12.Checked) {
                //代码走查时间
                TimeStamp.dmzc_time = this.dateTimePicker6.Value.ToShortDateString();
                TimeStamp.dmzc_format_time = this.dateTimePicker6.Value.ToLongDateString();
                TimeStamp.dmzcqr_time = this.dateTimePicker4.Value.ToShortDateString();
                TimeStamp.dmzcqr_format_time = this.dateTimePicker4.Value.ToLongDateString();
                TimeStamp.dmzchg_time = this.dateTimePicker5.Value.ToShortDateString();
                TimeStamp.dmzchg_format_time = this.dateTimePicker5.Value.ToLongDateString();
            }
            else
                ContentFlags.daimazoucha = 0;
            if(this.checkBox14.Checked) {
                //逻辑测试时间
                TimeStamp.ljcs_time = this.ljcs_shijian.Value.ToShortDateString();
                TimeStamp.ljcs_format_time = this.ljcs_shijian.Value.ToLongDateString();
            }
            else
                ContentFlags.luojiceshi = 0;
            if(this.xt_ceshi.Checked){
                //系统首轮测试时间
                TimeStamp.slxtcs_time = this.xtcs_shoulunshijian.Value.ToShortDateString();
                TimeStamp.slxtcs_format_time = this.xtcs_shoulunshijian.Value.ToLongDateString();
            }
            if(this.xt_ceshi.Checked && this.checkBox8.Checked) {
                //系统测试回归时间
                TimeStamp.xthgcs_time = this.xtcs_huiguishijian.Value.ToShortDateString();
                TimeStamp.xthgcs_format_time = this.xtcs_huiguishijian.Value.ToLongDateString();
            }
            else
                ContentFlags.xitonghuiguiceshi = 0;
            //动态配置项测试时间
            if(this.pzx_ceshi.Checked){
                TimeStamp.sldtcs_time = this.dtpz_shoulunshijian.Value.ToShortDateString();
                TimeStamp.sldtcs_format_time = this.dtpz_shoulunshijian.Value.ToLongDateString();
                TimeStamp.hgdtcs_time = this.dtpz_huigui.Value.ToShortDateString();
                TimeStamp.hgdtcs_format_time = this.dtpz_huigui.Value.ToLongDateString();
            }
            //测试报告评审时间
            TimeStamp.csbgps_time = this.csbg_pingshenshijian.Value.ToShortDateString();
            TimeStamp.csbgps_format_time = this.csbg_pingshenshijian.Value.ToLongDateString();

        }

        private void dtpz_huigui_ValueChanged(object sender, EventArgs e) {
            string dtpzhg_time = this.dtpz_huigui.Value.ToShortDateString();
            DateTime ljcs_time = DateHelper.cal_date(dtpzhg_time, -7);
            this.ljcs_shijian.Value = ljcs_time;
        }

       
        }
    }
