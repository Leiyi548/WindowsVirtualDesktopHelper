﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsVirtualDesktopHelper {
    public partial class AboutForm : Form {
        public AboutForm() {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e) {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            App.Instance.OpenURL("https://github.com/dankrusi/WindowsVirtualDesktopHelper");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            App.Instance.OpenURL("https://www.paypal.com/donate/?hosted_button_id=BG5FYMAHFG9V6");
        }

        private void AboutForm_Load(object sender, EventArgs e) {
            labelVersion.Text = "version "
                + Assembly.GetExecutingAssembly().GetName().Version.Major
                + "."
                + Assembly.GetExecutingAssembly().GetName().Version.Minor;
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            App.Instance.OpenURL("mailto:dan@dankrusi.com");
        }
    }
}
