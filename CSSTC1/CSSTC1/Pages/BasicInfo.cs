﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CSSTC1.InputProcessors;

namespace CSSTC1.Pages {
    public partial class BasicInfo : Form {
        public BasicInfo() {
            InitializeComponent();
            }
        public string Xm_biaoshi;
        public string Xm_mingcheng;
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

        BasicInfoProcessor processor = new BasicInfoProcessor();
        private void button1_Click(object sender, EventArgs e) {
            this.Xm_biaoshi = xm_biaoshi.Text;
            this.Xm_mingcheng = xm_mingcheng.Text;
            this.Pzx_ceshi = pzx_ceshi.Checked;
            this.Xt_ceshi = xt_ceshi.Checked;
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
            this.Xm_kaishishijian = xm_kaishishijian.Text;
            this.Xm_jieshushijian = xm_jieshushijian.Text;
            string[] values = { Xm_biaoshi, Xm_mingcheng, Pz_guanliyuan, Bcj_guanliyuan, Zlbz_renyuan, Cs_jianduyuan,                                     Sb_guanliyuan, Fuzhurren, Cw_fuzhuren, Zhuren, Wt_xingming, Wt_dianhua, Kf_xingming,                                      Kf_danwei, Kf_danwei, Kf_danwei, Xm_kaishishijian, Xm_jieshushijian};
            string[] bookmarks = { "Xm_biaoshi", "Xm_mingcheng", "Pz_guanliyuan", "Bcj_guanliyuan", "Zlbz_renyuan",                                          "Cs_jianduyuan", "Sb_guanliyuan", "Fuzhurren", "Cw_fuzhuren", "Zhuren", "Wt_xingming",                                    "Wt_dianhua", "Kf_xingming", "Kf_danwei", "Kf_danwei", "Kf_danwei", "Xm_kaishishijian",                                   "Xm_jieshushijian"};
            processor.fill_basic_info(bookmarks, values);
            }

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
            this.Csdg_pingshenshijian = csdg_pingshenshijian.Text;
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


        }
    }
