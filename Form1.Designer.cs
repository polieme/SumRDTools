﻿namespace SumRDTools
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.folderPathText = new Sunny.UI.UITextBox();
            this.chooseFolderBtn = new Sunny.UI.UIButton();
            this.summaryBtn = new Sunny.UI.UIButton();
            this.logTextBox = new Sunny.UI.UITextBox();
            this.uiStyleManager = new Sunny.UI.UIStyleManager(this.components);
            this.errorLogRichTextBox = new Sunny.UI.UIRichTextBox();
            this.SuspendLayout();
            // 
            // folderPathText
            // 
            this.folderPathText.ButtonFillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.folderPathText.ButtonFillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.folderPathText.ButtonFillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.folderPathText.ButtonRectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.folderPathText.ButtonRectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.folderPathText.ButtonRectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.folderPathText.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.folderPathText.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(248)))), ((int)(((byte)(248)))));
            this.folderPathText.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.folderPathText.Location = new System.Drawing.Point(18, 57);
            this.folderPathText.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.folderPathText.MinimumSize = new System.Drawing.Size(1, 16);
            this.folderPathText.Name = "folderPathText";
            this.folderPathText.Padding = new System.Windows.Forms.Padding(5);
            this.folderPathText.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.folderPathText.ScrollBarColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.folderPathText.ShowText = false;
            this.folderPathText.Size = new System.Drawing.Size(930, 35);
            this.folderPathText.Style = Sunny.UI.UIStyle.Gray;
            this.folderPathText.TabIndex = 0;
            this.folderPathText.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.folderPathText.Watermark = "";
            // 
            // chooseFolderBtn
            // 
            this.chooseFolderBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.chooseFolderBtn.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.chooseFolderBtn.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.chooseFolderBtn.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.chooseFolderBtn.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.chooseFolderBtn.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.chooseFolderBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.chooseFolderBtn.Location = new System.Drawing.Point(955, 57);
            this.chooseFolderBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.chooseFolderBtn.Name = "chooseFolderBtn";
            this.chooseFolderBtn.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.chooseFolderBtn.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.chooseFolderBtn.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.chooseFolderBtn.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.chooseFolderBtn.Size = new System.Drawing.Size(122, 35);
            this.chooseFolderBtn.Style = Sunny.UI.UIStyle.Gray;
            this.chooseFolderBtn.TabIndex = 1;
            this.chooseFolderBtn.Text = "选择文件夹...";
            this.chooseFolderBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.chooseFolderBtn.Click += new System.EventHandler(this.chooseFolderBtn_Click);
            // 
            // summaryBtn
            // 
            this.summaryBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.summaryBtn.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.summaryBtn.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.summaryBtn.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.summaryBtn.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.summaryBtn.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.summaryBtn.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.summaryBtn.Location = new System.Drawing.Point(1083, 57);
            this.summaryBtn.MinimumSize = new System.Drawing.Size(1, 1);
            this.summaryBtn.Name = "summaryBtn";
            this.summaryBtn.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.summaryBtn.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.summaryBtn.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.summaryBtn.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.summaryBtn.Size = new System.Drawing.Size(83, 35);
            this.summaryBtn.Style = Sunny.UI.UIStyle.Gray;
            this.summaryBtn.TabIndex = 2;
            this.summaryBtn.Text = "汇总";
            this.summaryBtn.TipsFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.summaryBtn.Click += new System.EventHandler(this.summaryBtn_Click);
            // 
            // logTextBox
            // 
            this.logTextBox.ButtonFillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.logTextBox.ButtonFillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.logTextBox.ButtonFillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.logTextBox.ButtonRectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.logTextBox.ButtonRectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.logTextBox.ButtonRectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.logTextBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.logTextBox.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(248)))), ((int)(((byte)(248)))));
            this.logTextBox.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.logTextBox.Location = new System.Drawing.Point(18, 101);
            this.logTextBox.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.logTextBox.MinimumSize = new System.Drawing.Size(1, 16);
            this.logTextBox.Multiline = true;
            this.logTextBox.Name = "logTextBox";
            this.logTextBox.Padding = new System.Windows.Forms.Padding(5);
            this.logTextBox.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.logTextBox.ScrollBarColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.logTextBox.ShowScrollBar = true;
            this.logTextBox.ShowText = false;
            this.logTextBox.Size = new System.Drawing.Size(336, 518);
            this.logTextBox.Style = Sunny.UI.UIStyle.Gray;
            this.logTextBox.TabIndex = 3;
            this.logTextBox.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.logTextBox.Watermark = "";
            // 
            // uiStyleManager
            // 
            this.uiStyleManager.DPIScale = true;
            this.uiStyleManager.GlobalFont = true;
            this.uiStyleManager.GlobalFontName = "微软雅黑";
            this.uiStyleManager.Style = Sunny.UI.UIStyle.Gray;
            // 
            // errorLogRichTextBox
            // 
            this.errorLogRichTextBox.FillColor = System.Drawing.Color.White;
            this.errorLogRichTextBox.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(248)))), ((int)(((byte)(248)))));
            this.errorLogRichTextBox.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.errorLogRichTextBox.Location = new System.Drawing.Point(370, 102);
            this.errorLogRichTextBox.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.errorLogRichTextBox.MinimumSize = new System.Drawing.Size(1, 1);
            this.errorLogRichTextBox.Name = "errorLogRichTextBox";
            this.errorLogRichTextBox.Padding = new System.Windows.Forms.Padding(2);
            this.errorLogRichTextBox.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.errorLogRichTextBox.ScrollBarColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.errorLogRichTextBox.ShowText = false;
            this.errorLogRichTextBox.Size = new System.Drawing.Size(796, 517);
            this.errorLogRichTextBox.Style = Sunny.UI.UIStyle.Gray;
            this.errorLogRichTextBox.TabIndex = 3;
            this.errorLogRichTextBox.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(1179, 637);
            this.ControlBoxFillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.Controls.Add(this.errorLogRichTextBox);
            this.Controls.Add(this.logTextBox);
            this.Controls.Add(this.summaryBtn);
            this.Controls.Add(this.chooseFolderBtn);
            this.Controls.Add(this.folderPathText);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.Style = Sunny.UI.UIStyle.Gray;
            this.Text = "企业研发活动情况表汇总数据工具";
            this.TitleColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.ZoomScaleRect = new System.Drawing.Rectangle(15, 15, 960, 641);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private Sunny.UI.UITextBox folderPathText;
        private Sunny.UI.UIButton chooseFolderBtn;
        private Sunny.UI.UIButton summaryBtn;
        private Sunny.UI.UITextBox logTextBox;
        private Sunny.UI.UIStyleManager uiStyleManager;
        private Sunny.UI.UIRichTextBox errorLogRichTextBox;
    }
}

