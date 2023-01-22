using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using ComponentFactory.Krypton.Toolkit;

namespace Infinium
{
    partial class LightStartForm
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new Container();
            ComponentResourceManager resources = new ComponentResourceManager(typeof(LightStartForm));
            this.NavigateMenuButtonsPalette = new KryptonPalette(this.components);
            this.AnimateTimer = new Timer(this.components);
            this.kryptonBorderEdge3 = new KryptonBorderEdge();
            this.MenuCloseButton = new KryptonButton();
            this.MinimizeButton = new KryptonButton();
            this.UserPanel = new Panel();
            this.StandardButtonsPalette = new KryptonPalette(this.components);
            this.UserLabel = new Label();
            this.PhotoBox = new PhotoBox();
            this.LogoutButton = new KryptonButton();
            this.label1 = new Label();
            this.HeaderPanel = new KryptonPanel();
            this.MessagesButton = new KryptonButton();
            this.MainPanel = new Panel();
            this.InfiniumMinimizeList = new InfiniumMinimizeList();
            this.kryptonBorderEdge7 = new KryptonBorderEdge();
            this.InfiniumStartMenu = new InfiniumStartMenu();
            this.InfiniumNotifyList = new InfiniumNotifyList();
            this.DayLabel = new Label();
            this.DateLabel = new Label();
            this.InfiniumClock = new InfiniumClock();
            this.TimeLabel = new Label();
            this.kryptonBorderEdge6 = new KryptonBorderEdge();
            this.kryptonBorderEdge5 = new KryptonBorderEdge();
            this.kryptonBorderEdge4 = new KryptonBorderEdge();
            this.kryptonBorderEdge2 = new KryptonBorderEdge();
            this.InfiniumTilesContainer = new InfiniumTilesContainer();
            this.BottomPanel = new KryptonPanel();
            this.label2 = new Label();
            this.CurrentTimer = new Timer(this.components);
            this.SplashTimer = new Timer(this.components);
            this.TileContextMenu = new KryptonContextMenu();
            this.kryptonContextMenuItems1 = new KryptonContextMenuItems();
            this.MenuAddToFavorite = new KryptonContextMenuItem();
            this.MenuRemoveFromFavorite = new KryptonContextMenuItem();
            this.ContextMenuPalette = new KryptonPalette(this.components);
            this.UserPanel.SuspendLayout();
            ((ISupportInitialize)(this.PhotoBox)).BeginInit();
            ((ISupportInitialize)(this.HeaderPanel)).BeginInit();
            this.HeaderPanel.SuspendLayout();
            this.MainPanel.SuspendLayout();
            ((ISupportInitialize)(this.BottomPanel)).BeginInit();
            this.BottomPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // NavigateMenuButtonsPalette
            // 
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Back.Color1 = Color.Transparent;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Back.Color2 = Color.Transparent;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Back.ColorStyle = PaletteColorStyle.Solid;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.Color1 = Color.DimGray;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.Color2 = Color.DimGray;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.Draw = InheritBool.False;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.DrawBorders = ((PaletteDrawBorders)((((PaletteDrawBorders.Top | PaletteDrawBorders.Bottom) 
            | PaletteDrawBorders.Left) 
            | PaletteDrawBorders.Right)));
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.Rounding = 0;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Back.Color1 = Color.Transparent;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Back.Color2 = Color.Transparent;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Back.ColorStyle = PaletteColorStyle.Solid;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.Color1 = Color.Gray;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.Color2 = Color.DimGray;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.Draw = InheritBool.False;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.DrawBorders = ((PaletteDrawBorders)((((PaletteDrawBorders.Top | PaletteDrawBorders.Bottom) 
            | PaletteDrawBorders.Left) 
            | PaletteDrawBorders.Right)));
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.Rounding = 0;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StatePressed.Back.Color1 = Color.Transparent;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StatePressed.Back.Color2 = Color.Transparent;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StatePressed.Back.ColorStyle = PaletteColorStyle.Solid;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StatePressed.Border.Color1 = Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StatePressed.Border.Color2 = Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StatePressed.Border.Draw = InheritBool.False;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StatePressed.Border.DrawBorders = ((PaletteDrawBorders)((((PaletteDrawBorders.Top | PaletteDrawBorders.Bottom) 
            | PaletteDrawBorders.Left) 
            | PaletteDrawBorders.Right)));
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StatePressed.Border.Rounding = 0;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StatePressed.Border.Width = 1;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Back.Color1 = Color.Transparent;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Back.Color2 = Color.Transparent;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Back.ColorStyle = PaletteColorStyle.Solid;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.Color1 = Color.Silver;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.Color2 = Color.Silver;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.Draw = InheritBool.True;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.DrawBorders = ((PaletteDrawBorders)((((PaletteDrawBorders.Top | PaletteDrawBorders.Bottom) 
            | PaletteDrawBorders.Left) 
            | PaletteDrawBorders.Right)));
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.Rounding = 0;
            this.NavigateMenuButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.Width = 1;
            // 
            // AnimateTimer
            // 
            this.AnimateTimer.Interval = 1;
            this.AnimateTimer.Tick += new EventHandler(this.AnimateTimer_Tick);
            // 
            // kryptonBorderEdge3
            // 
            this.kryptonBorderEdge3.Anchor = ((AnchorStyles)(((AnchorStyles.Bottom | AnchorStyles.Left) 
                                                              | AnchorStyles.Right)));
            this.kryptonBorderEdge3.AutoSizeMode = AutoSizeMode.GrowOnly;
            this.kryptonBorderEdge3.Location = new Point(22, 565);
            this.kryptonBorderEdge3.Name = "kryptonBorderEdge3";
            this.kryptonBorderEdge3.PaletteMode = PaletteMode.Office2010Black;
            this.kryptonBorderEdge3.Size = new Size(1225, 1);
            this.kryptonBorderEdge3.StateCommon.Color1 = Color.FromArgb(((int)(((byte)(180)))), ((int)(((byte)(180)))), ((int)(((byte)(180)))));
            this.kryptonBorderEdge3.StateCommon.Color2 = Color.Gray;
            this.kryptonBorderEdge3.Text = "kryptonBorderEdge3";
            // 
            // MenuCloseButton
            // 
            this.MenuCloseButton.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.MenuCloseButton.Location = new Point(1218, 5);
            this.MenuCloseButton.Name = "MenuCloseButton";
            this.MenuCloseButton.Palette = this.NavigateMenuButtonsPalette;
            this.MenuCloseButton.PaletteMode = PaletteMode.Custom;
            this.MenuCloseButton.Size = new Size(44, 40);
            this.MenuCloseButton.TabIndex = 13;
            this.MenuCloseButton.TabStop = false;
            this.MenuCloseButton.Values.Image = ((Image)(resources.GetObject("MenuCloseButton.Values.Image")));
            this.MenuCloseButton.Values.ImageTransparentColor = Color.Transparent;
            this.MenuCloseButton.Values.Text = "";
            this.MenuCloseButton.Click += new EventHandler(this.MenuCloseButton_Click_1);
            // 
            // MinimizeButton
            // 
            this.MinimizeButton.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.MinimizeButton.Location = new Point(1168, 5);
            this.MinimizeButton.Name = "MinimizeButton";
            this.MinimizeButton.Palette = this.NavigateMenuButtonsPalette;
            this.MinimizeButton.PaletteMode = PaletteMode.Custom;
            this.MinimizeButton.Size = new Size(44, 40);
            this.MinimizeButton.TabIndex = 23;
            this.MinimizeButton.TabStop = false;
            this.MinimizeButton.Values.Image = ((Image)(resources.GetObject("MinimizeButton.Values.Image")));
            this.MinimizeButton.Values.ImageTransparentColor = Color.Transparent;
            this.MinimizeButton.Values.Text = "";
            this.MinimizeButton.Click += new EventHandler(this.MinimizeButton_Click);
            // 
            // UserPanel
            // 
            this.UserPanel.Anchor = ((AnchorStyles)((AnchorStyles.Bottom | AnchorStyles.Right)));
            this.UserPanel.BackColor = Color.Transparent;
            this.UserPanel.Controls.Add(this.UserLabel);
            this.UserPanel.Controls.Add(this.PhotoBox);
            this.UserPanel.Controls.Add(this.LogoutButton);
            this.UserPanel.Location = new Point(990, 503);
            this.UserPanel.Name = "UserPanel";
            this.UserPanel.Size = new Size(276, 104);
            this.UserPanel.TabIndex = 19;
            // 
            // StandardButtonsPalette
            // 
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Back.Color1 = Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(184)))), ((int)(((byte)(238)))));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Back.Color2 = Color.FromArgb(((int)(((byte)(49)))), ((int)(((byte)(167)))), ((int)(((byte)(214)))));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Back.ColorStyle = PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.Color1 = Color.Transparent;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.Color2 = Color.Gray;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.ColorStyle = PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.DrawBorders = ((PaletteDrawBorders)((((PaletteDrawBorders.Top | PaletteDrawBorders.Bottom) 
            | PaletteDrawBorders.Left) 
            | PaletteDrawBorders.Right)));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.Rounding = 0;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Back.Color1 = Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(184)))), ((int)(((byte)(238)))));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Back.Color2 = Color.White;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Back.ColorStyle = PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.Color1 = Color.Transparent;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.Color2 = Color.Black;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.ColorStyle = PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.DrawBorders = ((PaletteDrawBorders)((((PaletteDrawBorders.Top | PaletteDrawBorders.Bottom) 
            | PaletteDrawBorders.Left) 
            | PaletteDrawBorders.Right)));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.Rounding = 0;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.Color1 = Color.White;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.Color2 = Color.White;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.ColorStyle = PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204)));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.MultiLine = InheritBool.True;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.MultiLineH = PaletteRelativeAlign.Center;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.TextV = PaletteRelativeAlign.Far;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Back.Color1 = Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(164)))), ((int)(((byte)(61)))));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Back.Color2 = Color.White;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Back.ColorStyle = PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.Color1 = Color.Transparent;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.Color2 = Color.Black;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.ColorStyle = PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.DrawBorders = ((PaletteDrawBorders)((((PaletteDrawBorders.Top | PaletteDrawBorders.Bottom) 
            | PaletteDrawBorders.Left) 
            | PaletteDrawBorders.Right)));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.Rounding = 0;
            // 
            // UserLabel
            // 
            this.UserLabel.AutoSize = true;
            this.UserLabel.BackColor = Color.Transparent;
            this.UserLabel.Font = new Font("Segoe UI", 14F, FontStyle.Regular, GraphicsUnit.Pixel);
            this.UserLabel.ForeColor = Color.DimGray;
            this.UserLabel.Location = new Point(110, 4);
            this.UserLabel.Name = "UserLabel";
            this.UserLabel.Size = new Size(124, 38);
            this.UserLabel.TabIndex = 17;
            this.UserLabel.Text = "Романчук\r\nАндрей Иванович";
            // 
            // PhotoBox
            // 
            this.PhotoBox.BackColor = Color.Transparent;
            this.PhotoBox.BorderColorCommon = Color.White;
            this.PhotoBox.BorderColorTracking = Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(164)))), ((int)(((byte)(61)))));
            this.PhotoBox.Cursor = Cursors.Hand;
            this.PhotoBox.DrawBorder = true;
            this.PhotoBox.Location = new Point(7, 4);
            this.PhotoBox.Name = "PhotoBox";
            this.PhotoBox.Size = new Size(100, 93);
            this.PhotoBox.SizeMode = PictureBoxSizeMode.StretchImage;
            this.PhotoBox.TabIndex = 185;
            this.PhotoBox.TabStop = false;
            this.PhotoBox.Click += new EventHandler(this.PhotoBox_Click);
            // 
            // LogoutButton
            // 
            this.LogoutButton.Location = new Point(115, 71);
            this.LogoutButton.Name = "LogoutButton";
            this.LogoutButton.Palette = this.StandardButtonsPalette;
            this.LogoutButton.PaletteMode = PaletteMode.Custom;
            this.LogoutButton.Size = new Size(90, 26);
            this.LogoutButton.StateCommon.Content.ShortText.Font = new Font("Segoe UI", 15F, FontStyle.Regular, GraphicsUnit.Pixel);
            this.LogoutButton.TabIndex = 18;
            this.LogoutButton.Values.Text = "Выход";
            this.LogoutButton.Click += new EventHandler(this.LogoutButton_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((AnchorStyles)(((AnchorStyles.Top | AnchorStyles.Left) 
                                                  | AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.BackColor = Color.Transparent;
            this.label1.Font = new Font("Segoe UI", 30F, FontStyle.Regular, GraphicsUnit.Pixel);
            this.label1.ForeColor = Color.White;
            this.label1.Location = new Point(3, 5);
            this.label1.Name = "label1";
            this.label1.Size = new Size(126, 41);
            this.label1.TabIndex = 70;
            this.label1.Text = "Infinium";
            this.label1.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // HeaderPanel
            // 
            this.HeaderPanel.Controls.Add(this.MessagesButton);
            this.HeaderPanel.Controls.Add(this.label1);
            this.HeaderPanel.Controls.Add(this.MinimizeButton);
            this.HeaderPanel.Controls.Add(this.MenuCloseButton);
            this.HeaderPanel.Dock = DockStyle.Top;
            this.HeaderPanel.Location = new Point(0, 0);
            this.HeaderPanel.Name = "HeaderPanel";
            this.HeaderPanel.Size = new Size(1270, 50);
            this.HeaderPanel.StateCommon.Color1 = Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.HeaderPanel.StateCommon.Color2 = Color.FromArgb(((int)(((byte)(49)))), ((int)(((byte)(49)))), ((int)(((byte)(49)))));
            this.HeaderPanel.StateCommon.ColorStyle = PaletteColorStyle.Linear;
            this.HeaderPanel.TabIndex = 307;
            // 
            // MessagesButton
            // 
            this.MessagesButton.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.MessagesButton.Location = new Point(1114, 6);
            this.MessagesButton.Name = "MessagesButton";
            this.MessagesButton.Palette = this.NavigateMenuButtonsPalette;
            this.MessagesButton.PaletteMode = PaletteMode.Custom;
            this.MessagesButton.Size = new Size(44, 40);
            this.MessagesButton.StateCommon.Content.Image.ImageV = PaletteRelativeAlign.Far;
            this.MessagesButton.TabIndex = 71;
            this.MessagesButton.TabStop = false;
            this.MessagesButton.Values.Image = ((Image)(resources.GetObject("MessagesButton.Values.Image")));
            this.MessagesButton.Values.ImageTransparentColor = Color.Transparent;
            this.MessagesButton.Values.Text = "";
            this.MessagesButton.Click += new EventHandler(this.MessagesButton_Click);
            // 
            // MainPanel
            // 
            this.MainPanel.Anchor = ((AnchorStyles)((((AnchorStyles.Top | AnchorStyles.Bottom) 
                                                      | AnchorStyles.Left) 
                                                     | AnchorStyles.Right)));
            this.MainPanel.BackColor = Color.White;
            this.MainPanel.Controls.Add(this.InfiniumMinimizeList);
            this.MainPanel.Controls.Add(this.kryptonBorderEdge7);
            this.MainPanel.Controls.Add(this.InfiniumStartMenu);
            this.MainPanel.Controls.Add(this.InfiniumNotifyList);
            this.MainPanel.Controls.Add(this.DayLabel);
            this.MainPanel.Controls.Add(this.DateLabel);
            this.MainPanel.Controls.Add(this.InfiniumClock);
            this.MainPanel.Controls.Add(this.TimeLabel);
            this.MainPanel.Controls.Add(this.kryptonBorderEdge6);
            this.MainPanel.Controls.Add(this.kryptonBorderEdge5);
            this.MainPanel.Controls.Add(this.UserPanel);
            this.MainPanel.Controls.Add(this.kryptonBorderEdge4);
            this.MainPanel.Controls.Add(this.kryptonBorderEdge2);
            this.MainPanel.Controls.Add(this.InfiniumTilesContainer);
            this.MainPanel.Location = new Point(0, 50);
            this.MainPanel.Name = "MainPanel";
            this.MainPanel.Size = new Size(1270, 620);
            this.MainPanel.TabIndex = 0;
            // 
            // InfiniumMinimizeList
            // 
            this.InfiniumMinimizeList.Anchor = ((AnchorStyles)(((AnchorStyles.Bottom | AnchorStyles.Left) 
                                                                | AnchorStyles.Right)));
            this.InfiniumMinimizeList.BackColor = Color.Transparent;
            // 
            // 
            // 
            this.InfiniumMinimizeList.HorizontalScrollBar.Anchor = ((AnchorStyles)(((AnchorStyles.Bottom | AnchorStyles.Left) 
            | AnchorStyles.Right)));
            this.InfiniumMinimizeList.HorizontalScrollBar.HorizontalScrollCommonShaftBackColor = Color.White;
            this.InfiniumMinimizeList.HorizontalScrollBar.HorizontalScrollCommonThumbButtonColor = Color.DarkGray;
            this.InfiniumMinimizeList.HorizontalScrollBar.HorizontalScrollTrackingShaftBackColor = Color.Gainsboro;
            this.InfiniumMinimizeList.HorizontalScrollBar.HorizontalScrollTrackingThumbButtonColor = Color.Gray;
            this.InfiniumMinimizeList.HorizontalScrollBar.Location = new Point(0, 81);
            this.InfiniumMinimizeList.HorizontalScrollBar.Name = "";
            this.InfiniumMinimizeList.HorizontalScrollBar.Offset = 0;
            this.InfiniumMinimizeList.HorizontalScrollBar.ScrollWheelOffset = 80;
            this.InfiniumMinimizeList.HorizontalScrollBar.Size = new Size(653, 12);
            this.InfiniumMinimizeList.HorizontalScrollBar.TabIndex = 0;
            this.InfiniumMinimizeList.HorizontalScrollBar.TotalControlWidth = 653;
            this.InfiniumMinimizeList.HorizontalScrollBar.Visible = false;
            this.InfiniumMinimizeList.Location = new Point(309, 507);
            this.InfiniumMinimizeList.Name = "InfiniumMinimizeList";
            this.InfiniumMinimizeList.Size = new Size(653, 93);
            this.InfiniumMinimizeList.TabIndex = 352;
            this.InfiniumMinimizeList.Text = "infiniumMinimizeList1";
            this.InfiniumMinimizeList.ItemClicked += new InfiniumMinimizeList.ItemClickedEventHandler(this.InfiniumMinimizeList_ItemClicked);
            this.InfiniumMinimizeList.CloseClicked += new InfiniumMinimizeList.ItemClickedEventHandler(this.InfiniumMinimizeList_CloseClicked);
            // 
            // kryptonBorderEdge7
            // 
            this.kryptonBorderEdge7.Anchor = ((AnchorStyles)(((AnchorStyles.Bottom | AnchorStyles.Left) 
                                                              | AnchorStyles.Right)));
            this.kryptonBorderEdge7.AutoSize = false;
            this.kryptonBorderEdge7.Location = new Point(295, 496);
            this.kryptonBorderEdge7.Name = "kryptonBorderEdge7";
            this.kryptonBorderEdge7.Size = new Size(680, 1);
            this.kryptonBorderEdge7.StateCommon.Color1 = Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(189)))), ((int)(((byte)(189)))));
            this.kryptonBorderEdge7.Text = "kryptonBorderEdge7";
            // 
            // InfiniumStartMenu
            // 
            this.InfiniumStartMenu.Anchor = ((AnchorStyles)(((AnchorStyles.Top | AnchorStyles.Bottom) 
                                                             | AnchorStyles.Left)));
            this.InfiniumStartMenu.BackColor = Color.Transparent;
            this.InfiniumStartMenu.ItemsDataTable = null;
            this.InfiniumStartMenu.Location = new Point(28, 20);
            this.InfiniumStartMenu.Name = "InfiniumStartMenu";
            this.InfiniumStartMenu.Selected = -1;
            this.InfiniumStartMenu.Size = new Size(241, 565);
            this.InfiniumStartMenu.TabIndex = 343;
            this.InfiniumStartMenu.Text = "infiniumStartMenu1";
            // 
            // 
            // 
            this.InfiniumStartMenu.VerticalScrollBar.Location = new Point(241, 0);
            this.InfiniumStartMenu.VerticalScrollBar.Name = "";
            this.InfiniumStartMenu.VerticalScrollBar.Offset = 0;
            this.InfiniumStartMenu.VerticalScrollBar.ScrollWheelOffset = 30;
            this.InfiniumStartMenu.VerticalScrollBar.TabIndex = 0;
            this.InfiniumStartMenu.VerticalScrollBar.TotalControlHeight = 565;
            this.InfiniumStartMenu.VerticalScrollBar.VerticalScrollCommonShaftBackColor = Color.LightBlue;
            this.InfiniumStartMenu.VerticalScrollBar.VerticalScrollCommonThumbButtonColor = Color.DarkGray;
            this.InfiniumStartMenu.VerticalScrollBar.VerticalScrollTrackingShaftBackColor = Color.Blue;
            this.InfiniumStartMenu.VerticalScrollBar.VerticalScrollTrackingThumbButtonColor = Color.Gray;
            this.InfiniumStartMenu.ItemClicked += new InfiniumStartMenu.ItemClickedEventHandler(this.InfiniumStartMenu_ItemClicked);
            // 
            // InfiniumNotifyList
            // 
            this.InfiniumNotifyList.Anchor = ((AnchorStyles)(((AnchorStyles.Top | AnchorStyles.Bottom) 
                                                              | AnchorStyles.Right)));
            this.InfiniumNotifyList.BackColor = Color.Transparent;
            this.InfiniumNotifyList.ItemsDataTable = null;
            this.InfiniumNotifyList.Location = new Point(990, 250);
            this.InfiniumNotifyList.Name = "InfiniumNotifyList";
            this.InfiniumNotifyList.Size = new Size(272, 230);
            this.InfiniumNotifyList.TabIndex = 338;
            this.InfiniumNotifyList.Text = "infiniumNotifyList1";
            // 
            // 
            // 
            this.InfiniumNotifyList.VerticalScrollBar.Anchor = ((AnchorStyles)(((AnchorStyles.Top | AnchorStyles.Bottom) 
            | AnchorStyles.Right)));
            this.InfiniumNotifyList.VerticalScrollBar.Location = new Point(260, 0);
            this.InfiniumNotifyList.VerticalScrollBar.Name = "";
            this.InfiniumNotifyList.VerticalScrollBar.Offset = 0;
            this.InfiniumNotifyList.VerticalScrollBar.ScrollWheelOffset = 30;
            this.InfiniumNotifyList.VerticalScrollBar.Size = new Size(12, 230);
            this.InfiniumNotifyList.VerticalScrollBar.TabIndex = 0;
            this.InfiniumNotifyList.VerticalScrollBar.TotalControlHeight = 230;
            this.InfiniumNotifyList.VerticalScrollBar.VerticalScrollCommonShaftBackColor = Color.White;
            this.InfiniumNotifyList.VerticalScrollBar.VerticalScrollCommonThumbButtonColor = Color.DarkGray;
            this.InfiniumNotifyList.VerticalScrollBar.VerticalScrollTrackingShaftBackColor = Color.Gainsboro;
            this.InfiniumNotifyList.VerticalScrollBar.VerticalScrollTrackingThumbButtonColor = Color.Gray;
            this.InfiniumNotifyList.VerticalScrollBar.Visible = false;
            this.InfiniumNotifyList.ItemClicked += new InfiniumNotifyList.ItemClickedEventHandler(this.InfiniumNotifyList_ItemClicked);
            // 
            // DayLabel
            // 
            this.DayLabel.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.DayLabel.BackColor = Color.Transparent;
            this.DayLabel.Font = new Font("Segoe UI", 15F, FontStyle.Regular, GraphicsUnit.Pixel);
            this.DayLabel.ForeColor = Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(170)))), ((int)(((byte)(56)))));
            this.DayLabel.Location = new Point(990, 201);
            this.DayLabel.Name = "DayLabel";
            this.DayLabel.Size = new Size(271, 23);
            this.DayLabel.TabIndex = 333;
            this.DayLabel.Text = "среда";
            this.DayLabel.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // DateLabel
            // 
            this.DateLabel.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.DateLabel.BackColor = Color.Transparent;
            this.DateLabel.Font = new Font("Segoe UI", 17F, FontStyle.Regular, GraphicsUnit.Pixel);
            this.DateLabel.ForeColor = Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(184)))), ((int)(((byte)(238)))));
            this.DateLabel.Location = new Point(990, 180);
            this.DateLabel.Name = "DateLabel";
            this.DateLabel.Size = new Size(271, 23);
            this.DateLabel.TabIndex = 332;
            this.DateLabel.Text = "05 февраля 2014";
            this.DateLabel.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // InfiniumClock
            // 
            this.InfiniumClock.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.InfiniumClock.Cursor = Cursors.Hand;
            this.InfiniumClock.Image = ((Image)(resources.GetObject("InfiniumClock.Image")));
            this.InfiniumClock.Location = new Point(1058, 17);
            this.InfiniumClock.MarginLines = 25;
            this.InfiniumClock.Name = "InfiniumClock";
            this.InfiniumClock.Size = new Size(140, 140);
            this.InfiniumClock.TabIndex = 326;
            this.InfiniumClock.Text = "infiniumClock1";
            this.InfiniumClock.Click += new EventHandler(this.InfiniumClock_Click);
            // 
            // TimeLabel
            // 
            this.TimeLabel.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.TimeLabel.BackColor = Color.Transparent;
            this.TimeLabel.Font = new Font("Segoe UI", 14F, FontStyle.Regular, GraphicsUnit.Pixel);
            this.TimeLabel.ForeColor = Color.DimGray;
            this.TimeLabel.Location = new Point(990, 160);
            this.TimeLabel.Name = "TimeLabel";
            this.TimeLabel.Size = new Size(271, 20);
            this.TimeLabel.TabIndex = 307;
            this.TimeLabel.Text = "22:49:14";
            this.TimeLabel.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // kryptonBorderEdge6
            // 
            this.kryptonBorderEdge6.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.kryptonBorderEdge6.AutoSize = false;
            this.kryptonBorderEdge6.Location = new Point(990, 225);
            this.kryptonBorderEdge6.Name = "kryptonBorderEdge6";
            this.kryptonBorderEdge6.Size = new Size(270, 1);
            this.kryptonBorderEdge6.StateCommon.Color1 = Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(189)))), ((int)(((byte)(189)))));
            this.kryptonBorderEdge6.Text = "kryptonBorderEdge6";
            // 
            // kryptonBorderEdge5
            // 
            this.kryptonBorderEdge5.Anchor = ((AnchorStyles)((AnchorStyles.Bottom | AnchorStyles.Right)));
            this.kryptonBorderEdge5.AutoSize = false;
            this.kryptonBorderEdge5.Location = new Point(990, 496);
            this.kryptonBorderEdge5.Name = "kryptonBorderEdge5";
            this.kryptonBorderEdge5.Size = new Size(270, 1);
            this.kryptonBorderEdge5.StateCommon.Color1 = Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(189)))), ((int)(((byte)(189)))));
            this.kryptonBorderEdge5.Text = "kryptonBorderEdge5";
            // 
            // kryptonBorderEdge4
            // 
            this.kryptonBorderEdge4.Anchor = ((AnchorStyles)(((AnchorStyles.Top | AnchorStyles.Bottom) 
                                                              | AnchorStyles.Right)));
            this.kryptonBorderEdge4.AutoSize = false;
            this.kryptonBorderEdge4.Location = new Point(982, 20);
            this.kryptonBorderEdge4.Name = "kryptonBorderEdge4";
            this.kryptonBorderEdge4.Orientation = Orientation.Vertical;
            this.kryptonBorderEdge4.Size = new Size(1, 580);
            this.kryptonBorderEdge4.StateCommon.Color1 = Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(189)))), ((int)(((byte)(189)))));
            this.kryptonBorderEdge4.Text = "kryptonBorderEdge4";
            // 
            // kryptonBorderEdge2
            // 
            this.kryptonBorderEdge2.Anchor = ((AnchorStyles)(((AnchorStyles.Top | AnchorStyles.Bottom) 
                                                              | AnchorStyles.Left)));
            this.kryptonBorderEdge2.AutoSize = false;
            this.kryptonBorderEdge2.Location = new Point(286, 20);
            this.kryptonBorderEdge2.Name = "kryptonBorderEdge2";
            this.kryptonBorderEdge2.Orientation = Orientation.Vertical;
            this.kryptonBorderEdge2.Size = new Size(1, 580);
            this.kryptonBorderEdge2.StateCommon.Color1 = Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(189)))), ((int)(((byte)(189)))));
            this.kryptonBorderEdge2.Text = "kryptonBorderEdge2";
            // 
            // InfiniumTilesContainer
            // 
            this.InfiniumTilesContainer.Anchor = ((AnchorStyles)((((AnchorStyles.Top | AnchorStyles.Bottom) 
                                                                   | AnchorStyles.Left) 
                                                                  | AnchorStyles.Right)));
            this.InfiniumTilesContainer.BackColor = Color.Transparent;
            this.InfiniumTilesContainer.ItemsDataTable = null;
            this.InfiniumTilesContainer.Location = new Point(286, 20);
            this.InfiniumTilesContainer.Name = "InfiniumTilesContainer";
            this.InfiniumTilesContainer.Size = new Size(697, 460);
            this.InfiniumTilesContainer.TabIndex = 297;
            this.InfiniumTilesContainer.Text = "infiniumTilesContainer1";
            // 
            // 
            // 
            this.InfiniumTilesContainer.VerticalScrollBar.Anchor = ((AnchorStyles)(((AnchorStyles.Top | AnchorStyles.Bottom) 
            | AnchorStyles.Right)));
            this.InfiniumTilesContainer.VerticalScrollBar.Location = new Point(685, 0);
            this.InfiniumTilesContainer.VerticalScrollBar.Name = "";
            this.InfiniumTilesContainer.VerticalScrollBar.Offset = 0;
            this.InfiniumTilesContainer.VerticalScrollBar.ScrollWheelOffset = 100;
            this.InfiniumTilesContainer.VerticalScrollBar.Size = new Size(12, 460);
            this.InfiniumTilesContainer.VerticalScrollBar.TabIndex = 0;
            this.InfiniumTilesContainer.VerticalScrollBar.TotalControlHeight = 460;
            this.InfiniumTilesContainer.VerticalScrollBar.VerticalScrollCommonShaftBackColor = Color.Transparent;
            this.InfiniumTilesContainer.VerticalScrollBar.VerticalScrollCommonThumbButtonColor = Color.DarkGray;
            this.InfiniumTilesContainer.VerticalScrollBar.VerticalScrollTrackingShaftBackColor = Color.Transparent;
            this.InfiniumTilesContainer.VerticalScrollBar.VerticalScrollTrackingThumbButtonColor = Color.Gray;
            this.InfiniumTilesContainer.VerticalScrollBar.Visible = false;
            this.InfiniumTilesContainer.ItemClicked += new InfiniumTilesContainer.ItemClickedEventHandler(this.InfiniumTilesContainer_ItemClicked);
            this.InfiniumTilesContainer.RightMouseClicked += new InfiniumTilesContainer.ItemClickedEventHandler(this.InfiniumTilesContainer_RightMouseClicked);
            // 
            // BottomPanel
            // 
            this.BottomPanel.Controls.Add(this.label2);
            this.BottomPanel.Dock = DockStyle.Bottom;
            this.BottomPanel.Location = new Point(0, 670);
            this.BottomPanel.Name = "BottomPanel";
            this.BottomPanel.Size = new Size(1270, 50);
            this.BottomPanel.StateCommon.Color1 = Color.FromArgb(((int)(((byte)(231)))), ((int)(((byte)(240)))), ((int)(((byte)(247)))));
            this.BottomPanel.StateCommon.ColorStyle = PaletteColorStyle.Solid;
            this.BottomPanel.TabIndex = 309;
            // 
            // label2
            // 
            this.label2.Anchor = AnchorStyles.None;
            this.label2.AutoSize = true;
            this.label2.BackColor = Color.Transparent;
            this.label2.Font = new Font("Segoe UI", 14F, FontStyle.Regular, GraphicsUnit.Pixel);
            this.label2.ForeColor = Color.DarkGray;
            this.label2.Location = new Point(574, 15);
            this.label2.Name = "label2";
            this.label2.Size = new Size(111, 19);
            this.label2.TabIndex = 315;
            this.label2.Text = "обновлений нет";
            // 
            // CurrentTimer
            // 
            this.CurrentTimer.Enabled = true;
            this.CurrentTimer.Interval = 500;
            this.CurrentTimer.Tick += new EventHandler(this.CurrentTimer_Tick);
            // 
            // SplashTimer
            // 
            this.SplashTimer.Enabled = true;
            this.SplashTimer.Tick += new EventHandler(this.SplashTimer_Tick);
            // 
            // TileContextMenu
            // 
            this.TileContextMenu.Items.AddRange(new KryptonContextMenuItemBase[] {
            this.kryptonContextMenuItems1});
            this.TileContextMenu.Palette = this.ContextMenuPalette;
            // 
            // kryptonContextMenuItems1
            // 
            this.kryptonContextMenuItems1.Items.AddRange(new KryptonContextMenuItemBase[] {
            this.MenuAddToFavorite,
            this.MenuRemoveFromFavorite});
            // 
            // MenuAddToFavorite
            // 
            this.MenuAddToFavorite.Image = ((Image)(resources.GetObject("MenuAddToFavorite.Image")));
            this.MenuAddToFavorite.Text = "Добавить в мой список";
            this.MenuAddToFavorite.Click += new EventHandler(this.MenuAddToFavorite_Click);
            // 
            // MenuRemoveFromFavorite
            // 
            this.MenuRemoveFromFavorite.Text = "Удалить из списка";
            this.MenuRemoveFromFavorite.Visible = false;
            this.MenuRemoveFromFavorite.Click += new EventHandler(this.MenuRemoveFromFavorite_Click);
            // 
            // ContextMenuPalette
            // 
            this.ContextMenuPalette.ContextMenu.StateCommon.ControlOuter.Border.DrawBorders = ((PaletteDrawBorders)((((PaletteDrawBorders.Top | PaletteDrawBorders.Bottom) 
            | PaletteDrawBorders.Left) 
            | PaletteDrawBorders.Right)));
            this.ContextMenuPalette.ContextMenu.StateCommon.ControlOuter.Border.Rounding = 0;
            this.ContextMenuPalette.ContextMenu.StateCommon.ItemHighlight.Back.Color1 = Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(248)))), ((int)(((byte)(248)))));
            this.ContextMenuPalette.ContextMenu.StateCommon.ItemHighlight.Back.ColorStyle = PaletteColorStyle.Solid;
            this.ContextMenuPalette.ContextMenu.StateCommon.ItemHighlight.Border.Draw = InheritBool.False;
            this.ContextMenuPalette.ContextMenu.StateCommon.ItemHighlight.Border.DrawBorders = ((PaletteDrawBorders)((((PaletteDrawBorders.Top | PaletteDrawBorders.Bottom) 
            | PaletteDrawBorders.Left) 
            | PaletteDrawBorders.Right)));
            this.ContextMenuPalette.ContextMenu.StateCommon.ItemHighlight.Border.Rounding = 0;
            this.ContextMenuPalette.ContextMenu.StateCommon.ItemImageColumn.Border.Draw = InheritBool.False;
            this.ContextMenuPalette.ContextMenu.StateCommon.ItemImageColumn.Border.DrawBorders = ((PaletteDrawBorders)((((PaletteDrawBorders.Top | PaletteDrawBorders.Bottom) 
            | PaletteDrawBorders.Left) 
            | PaletteDrawBorders.Right)));
            this.ContextMenuPalette.ContextMenu.StateNormal.ItemTextStandard.LongText.Color1 = Color.Black;
            this.ContextMenuPalette.ContextMenu.StateNormal.ItemTextStandard.LongText.Font = new Font("Segoe UI", 16F, FontStyle.Regular, GraphicsUnit.Pixel);
            this.ContextMenuPalette.ContextMenu.StateNormal.ItemTextStandard.Padding = new Padding(-1, 6, -1, 6);
            // 
            // LightStartForm
            // 
            this.AutoScaleMode = AutoScaleMode.None;
            this.BackColor = Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.ClientSize = new Size(1270, 720);
            this.Controls.Add(this.MainPanel);
            this.Controls.Add(this.BottomPanel);
            this.Controls.Add(this.HeaderPanel);
            this.Controls.Add(this.kryptonBorderEdge3);
            this.Cursor = Cursors.Default;
            this.DoubleBuffered = true;
            this.FormBorderStyle = FormBorderStyle.None;
            this.Icon = ((Icon)(resources.GetObject("$this.Icon")));
            this.Name = "LightStartForm";
            this.Opacity = 0D;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Infinium";
            this.WindowState = FormWindowState.Maximized;
            this.Activated += new EventHandler(this.LightStartForm_Activated);
            this.FormClosing += new FormClosingEventHandler(this.LightStartForm_FormClosing);
            this.Load += new EventHandler(this.LightStartForm_Load);
            this.UserPanel.ResumeLayout(false);
            this.UserPanel.PerformLayout();
            ((ISupportInitialize)(this.PhotoBox)).EndInit();
            ((ISupportInitialize)(this.HeaderPanel)).EndInit();
            this.HeaderPanel.ResumeLayout(false);
            this.HeaderPanel.PerformLayout();
            this.MainPanel.ResumeLayout(false);
            ((ISupportInitialize)(this.BottomPanel)).EndInit();
            this.BottomPanel.ResumeLayout(false);
            this.BottomPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Timer AnimateTimer;
        private KryptonPalette NavigateMenuButtonsPalette;
        private KryptonButton MenuCloseButton;
        private KryptonBorderEdge kryptonBorderEdge3;
        private Label UserLabel;
        private KryptonButton LogoutButton;
        private Panel UserPanel;
        private KryptonButton MinimizeButton;
        private PhotoBox PhotoBox;
        private Label label1;
        private KryptonPanel HeaderPanel;
        private Panel MainPanel;
        private KryptonPanel BottomPanel;
        private InfiniumTilesContainer InfiniumTilesContainer;
        private KryptonBorderEdge kryptonBorderEdge4;
        private KryptonBorderEdge kryptonBorderEdge2;
        private KryptonPalette StandardButtonsPalette;
        private KryptonBorderEdge kryptonBorderEdge6;
        private KryptonBorderEdge kryptonBorderEdge5;
        private Label TimeLabel;
        private InfiniumClock InfiniumClock;
        private Timer CurrentTimer;
        private Label DayLabel;
        private Label DateLabel;
        private InfiniumNotifyList InfiniumNotifyList;
        private Label label2;
        private InfiniumStartMenu InfiniumStartMenu;
        private Timer SplashTimer;
        private KryptonContextMenu TileContextMenu;
        private KryptonContextMenuItems kryptonContextMenuItems1;
        private KryptonContextMenuItem MenuAddToFavorite;
        private KryptonPalette ContextMenuPalette;
        private KryptonContextMenuItem MenuRemoveFromFavorite;
        private InfiniumMinimizeList InfiniumMinimizeList;
        private KryptonBorderEdge kryptonBorderEdge7;
        private KryptonButton MessagesButton;
    }
}

