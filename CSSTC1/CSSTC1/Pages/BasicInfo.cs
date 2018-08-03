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
        public string Wd_shenchashijian;
        public string Wd_querenshijian;
        public string Wd_huiguishijian;
        public string Jtfx_shenchashijian;
        public string Jtfx_querenshijian;
        public string Jtfx_huiguishijian;
        public string Dmsc_shenchashijian;
        public string Dmsc_querenshijian;
        public string Dmsc_huiguishijian;
        public string Dmzc_shenchashijian;
        public string Dmzc_querenshijian;
        public string Dmzc_huiguishijian;
        public string Csdg_pingshenshijian;
        public string Dtpz_shoulunshijian;
        public string Dtpz_huiguishijian;
        public string Ljcs_shijian;
        public string Xtcs_shoulunshijian;
        public string Xtcs_huiguishijian;
        public string Csbg_pingshenshijian;
        public bool Yz_renwushu;
        public bool Js_fangan;
        public bool Js_guigeshu;
        public bool Rjyz_renwushu;
        public bool Rjxq_guigeshuoming;

        BasicInfoProcessor processor = new BasicInfoProcessor();

        private void wd_shenchashijian_ValueChanged(object sender, EventArgs e) {
            this.Wd_shenchashijian = wd_shenchashijian.Text;
            }

        private void wd_querenshijian_ValueChanged(object sender, EventArgs e) {
            this.Wd_querenshijian = wd_querenshijian.Text;
            }

        private void wd_huiguishijian_ValueChanged(object sender, EventArgs e) {
            this.Wd_huiguishijian = wd_huiguishijian.Text;
            }

        private void jtfx_shenchashijian_ValueChanged(object sender, EventArgs e) {
            this.Jtfx_shenchashijian = jtfx_shenchashijian.Text;
        }

        private void jtfx_querenshijian_ValueChanged(object sender, EventArgs e) {
            this.Jtfx_querenshijian = jtfx_querenshijian.Text;
        }

        private void jt_huiguishijian_ValueChanged(object sender, EventArgs e) {
            this.Jtfx_huiguishijian = jt_huiguishijian.Text;
        }

        private void dmsc_shijian_ValueChanged(object sender, EventArgs e) {
            this.Dmsc_shenchashijian = dmsc_shijian.Text;
        }

        private void dmsc_querenshijian_ValueChanged(object sender, EventArgs e) {
            this.Dmsc_querenshijian = dmsc_querenshijian.Text;
        }

        private void dmhg_huiguishijian_ValueChanged(object sender, EventArgs e) {
            this.Dmzc_huiguishijian = dmhg_huiguishijian.Text;
        }

        private void dmzc_shenchashijian_ValueChanged(object sender, EventArgs e) {
            this.Dmzc_shenchashijian = dmzc_shenchashijian.Text;
        }

        private void dmzc_querenshijian_ValueChanged(object sender, EventArgs e) {
            this.Dmzc_querenshijian = dmzc_querenshijian.Text;
        }

        private void dmzc_huiguishijian_ValueChanged(object sender, EventArgs e) {
            this.Dmzc_huiguishijian = dmzc_huiguishijian.Text;
        }

        private void csdg_pingshenshijian_ValueChanged(object sender, EventArgs e) {
            //this.Csdg_pingshenshijian = csdg_pingshenshijian.Text;
            //测试大纲评审时间
            TimeStamp.ceshisc_time = csdg_pingshenshijian.Text;
            //测试大纲评审时间文字形式
            TimeStamp.ceshisc_format_time = csdg_pingshenshijian.Value.ToLongDateString();
        }

        private void csxq_pingshenshijian_ValueChanged(object sender, EventArgs e) {
            //测试大纲或测试说明评审时间
            TimeStamp.ceshisc_time = csxq_pingshenshijian.Text;
            //测试大纲或测试说明评审时间文字形式
            TimeStamp.ceshisc_format_time = csxq_pingshenshijian.Value.ToLongDateString();
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e) {
            this.xtcs_huiguishijian.Enabled = this.checkBox8.Checked;
        }

        private void dtpz_shoulunshijian_ValueChanged(object sender, EventArgs e) {
            this.Dtpz_shoulunshijian = dtpz_shoulunshijian.Text;
        }

        private void dtpz_huigui_ValueChanged(object sender, EventArgs e) {
            this.Dtpz_huiguishijian = dtpz_huigui.Text;
        }

        private void ljcs_shijian_ValueChanged(object sender, EventArgs e) {
            this.Ljcs_shijian = ljcs_shijian.Text;
        }

        private void xtcs_shoulunshijian_ValueChanged(object sender, EventArgs e) {
            this.Xtcs_shoulunshijian = xtcs_shoulunshijian.Text;
        }

        private void xtcs_huiguishijian_ValueChanged(object sender, EventArgs e) {
            this.Xtcs_huiguishijian = xtcs_huiguishijian.Text;
        }

        private void csbg_pingshenshijian_ValueChanged(object sender, EventArgs e) {
            this.Csbg_pingshenshijian = csbg_pingshenshijian.Text;
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
        #endregion

        //提交信息
        private void button1_Click(object sender, EventArgs e) {
            button1.Enabled = false;
            this.Xm_biaoshi = xm_biaoshi.Text;
            this.Xm_mingcheng = xm_mingcheng.Text;
            this.Rj_mingcheng = rj_mingcheng.Text;
            ContentFlags.peizhiceshi = pzx_ceshi.Checked;
            ContentFlags.xitongceshi = xt_ceshi.Checked;
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
            TimeStamp.xiangmukaishishijian = xm_kaishishijian.Text;
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
                CommonUtils.OperationHelper.delete_section(FilePaths.save_root_file, bookmarks);
            }
            else{
                string[] bookmarks = { "测试需求1", "测试需求2", "测试需求3", "测试需求4", 
                                         "测试需求5", "测试需求6", "测试需求7"};
                CommonUtils.OperationHelper.delete_section(FilePaths.save_root_file, bookmarks);
            }
        }

        }
    }
