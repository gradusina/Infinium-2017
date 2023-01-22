using Infinium.Modules.Marketing.Expedition;

using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class CatalogForm : Form
    {
        private const int EHide = 2;
        private const int EShow = 1;
        private const int EClose = 3;
        private const int EMainMenu = 4;
        private const int IEditor = 93;

        private int _formEvent;

        private bool _bCanEdit;
        private bool _needSplash;

        private LightStartForm _lightStartForm;

        private Form _topForm;
        private Infinium.FrontsCatalog _frontsCatalog;
        private Infinium.DecorCatalog _decorCatalog;
        private CabFurLabel _cabFurLabelManager;
        private CubeLabel _cubeLabelManager;
        private SampleLabel _sampleLabelManager;
        private DataTable _attachmentsDt;
        private DataTable _frontsDt;
        private DataTable _decorDt;
        private DataTable _cabFurDt;

        public CatalogForm(LightStartForm tLightStartForm)
        {
            InitializeComponent();
            _lightStartForm = tLightStartForm;
            this.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size;

            Initialize();
            btnSetDescription.Enabled = false;
            kryptonButton1.Enabled = false;
            SetPermissions();

            while (!SplashForm.bCreated) ;
        }

        private void SetPermissions()
        {
            DataTable dt = RolesAndPermissionsManager.GetPermissions(Security.CurrentUserID);

            if (dt.Rows.Count == 0)
                return;

            if (!dt.Select("RoleID = 12").Any())
                return;
            _bCanEdit = true;
            btnSetDescription.Enabled = true;
            kryptonButton1.Enabled = true;
            //kryptonContextMenuItem7.Enabled = true;
            //kryptonContextMenuItem9.Enabled = true;
            //kryptonContextMenuItem10.Enabled = true;
            //kryptonContextMenuItem11.Enabled = true;
        }

        private void CatalogForm_Shown(object sender, EventArgs e)
        {
            FrontsCatalogPanel.BringToFront();
            while (!SplashForm.bCreated) ;
            _formEvent = EShow;
            AnimateTimer.Enabled = true;
            _needSplash = true;
        }

        private void AnimateTimer_Tick(object sender, EventArgs e)
        {
            if (!DatabaseConfigsManager.Animation)
            {
                this.Opacity = 1;

                if (_formEvent == EClose || _formEvent == EHide)
                {
                    AnimateTimer.Enabled = false;

                    if (_formEvent == EClose)
                    {

                        _lightStartForm.CloseForm(this);
                    }

                    if (_formEvent == EHide)
                    {
                        _needSplash = false;
                        _lightStartForm.HideForm(this);
                    }


                    return;
                }

                if (_formEvent == EShow)
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                    return;
                }

            }

            if (_formEvent == EClose || _formEvent == EHide)
            {
                if (Convert.ToDecimal(this.Opacity) != Convert.ToDecimal(0.00))
                    this.Opacity = Convert.ToDouble(Convert.ToDecimal(this.Opacity) - Convert.ToDecimal(0.05));
                else
                {
                    AnimateTimer.Enabled = false;

                    if (_formEvent == EClose)
                    {
                        _lightStartForm.CloseForm(this);
                    }

                    if (_formEvent == EHide)
                    {
                        _needSplash = false;
                        _lightStartForm.HideForm(this);
                    }

                }

                return;
            }


            if (_formEvent == EShow || _formEvent == EShow)
            {
                if (this.Opacity != 1)
                    this.Opacity += 0.05;
                else
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                }

                return;
            }
        }

        private void NavigateMenuCloseButton_Click(object sender, EventArgs e)
        {
            _formEvent = EClose;
            AnimateTimer.Enabled = true;
        }

/*
        private void NavigateMenuBackButton_Click(object sender, EventArgs e)
        {
            _formEvent = EHide;
            AnimateTimer.Enabled = true;
        }
*/

/*
        private void NavigateMenuHerculesButton_Click(object sender, EventArgs e)
        {
            _formEvent = EMainMenu;
            AnimateTimer.Enabled = true;
        }
*/


        private void Initialize()
        {
            _cabFurLabelManager = new CabFurLabel();
            _cubeLabelManager = new CubeLabel();
            _cabFurDt = new DataTable();
            _cabFurDt.Columns.Add(new DataColumn("TechStoreName", Type.GetType("System.String")));
            _cabFurDt.Columns.Add(new DataColumn("TechStoreSubGroupName", Type.GetType("System.String")));
            _cabFurDt.Columns.Add(new DataColumn("SubGroupNotes", Type.GetType("System.String")));
            _cabFurDt.Columns.Add(new DataColumn("SubGroupNotes1", Type.GetType("System.String")));
            _cabFurDt.Columns.Add(new DataColumn("SubGroupNotes2", Type.GetType("System.String")));
            _cabFurDt.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
            _cabFurDt.Columns.Add(new DataColumn("Color2", Type.GetType("System.String")));
            _cabFurDt.Columns.Add(new DataColumn("Height", Type.GetType("System.String")));
            _cabFurDt.Columns.Add(new DataColumn("Width", Type.GetType("System.String")));
            _cabFurDt.Columns.Add(new DataColumn("Length", Type.GetType("System.String")));
            _cabFurDt.Columns.Add(new DataColumn("PositionsCount", Type.GetType("System.String")));
            _cabFurDt.Columns.Add(new DataColumn("LabelsCount", Type.GetType("System.Int32")));
            _cabFurDt.Columns.Add(new DataColumn("FactoryType", Type.GetType("System.Int32")));
            _cabFurDt.Columns.Add(new DataColumn("DecorConfigID", Type.GetType("System.Int32")));

            _attachmentsDt = new DataTable();
            _attachmentsDt.Columns.Add(new DataColumn("FileName", Type.GetType("System.String")));
            _attachmentsDt.Columns.Add(new DataColumn("Extension", Type.GetType("System.String")));
            _attachmentsDt.Columns.Add(new DataColumn("Path", Type.GetType("System.String")));
            if (_frontsDt == null)
            {
                _frontsDt = new DataTable();
                _frontsDt.Columns.Add(new DataColumn("Front", Type.GetType("System.String")));
                _frontsDt.Columns.Add(new DataColumn("FrameColor", Type.GetType("System.String")));
                _frontsDt.Columns.Add(new DataColumn("InsetType", Type.GetType("System.String")));
                _frontsDt.Columns.Add(new DataColumn("InsetColor", Type.GetType("System.String")));
                _frontsDt.Columns.Add(new DataColumn("PositionsCount", Type.GetType("System.String")));
                _frontsDt.Columns.Add(new DataColumn("LabelsCount", Type.GetType("System.Int32")));
                _frontsDt.Columns.Add(new DataColumn("FactoryType", Type.GetType("System.Int32")));
                _frontsDt.Columns.Add(new DataColumn("FrontConfigID", Type.GetType("System.Int32")));
                _frontsDt.Columns.Add(new DataColumn("FrontID", Type.GetType("System.Int32")));
                _frontsDt.Columns.Add(new DataColumn("ColorID", Type.GetType("System.Int32")));
                _frontsDt.Columns.Add(new DataColumn("TechnoColorID", Type.GetType("System.Int32")));
            }
            if (_decorDt == null)
            {
                _decorDt = new DataTable();
                _decorDt.Columns.Add(new DataColumn("Product", Type.GetType("System.String")));
                _decorDt.Columns.Add(new DataColumn("Decor", Type.GetType("System.String")));
                _decorDt.Columns.Add(new DataColumn("Color", Type.GetType("System.String")));
                _decorDt.Columns.Add(new DataColumn("Length", Type.GetType("System.String")));
                _decorDt.Columns.Add(new DataColumn("Height", Type.GetType("System.String")));
                _decorDt.Columns.Add(new DataColumn("Width", Type.GetType("System.String")));
                _decorDt.Columns.Add(new DataColumn("PositionsCount", Type.GetType("System.String")));
                _decorDt.Columns.Add(new DataColumn("LabelsCount", Type.GetType("System.Int32")));
                _decorDt.Columns.Add(new DataColumn("FactoryType", Type.GetType("System.Int32")));
                _decorDt.Columns.Add(new DataColumn("DecorConfigID", Type.GetType("System.Int32")));
            }
            _sampleLabelManager = new SampleLabel();

            _frontsCatalog = new FrontsCatalog(1);
            _decorCatalog = new DecorCatalog(1);

            FilterClientsDataGrid.DataSource = _frontsCatalog.FilterClientsBindingSource;
            FilterClientsDataGrid.Columns["ClientID"].Visible = false;

            FrontsCatalogBinding();
            DecorCatalogBinding();
            FrontsGridsSettings();
            DecorGridsSettings();
        }

        #region FRONTS FUNCTIONS
        private void FrontsCatalogBinding()
        {
            FrontsDataGrid.SelectionChanged -= FrontsDataGrid_SelectionChanged;
            FrameColorsDataGrid.SelectionChanged -= FrameColorsDataGrid_SelectionChanged;
            TechnoFrameColorsDataGrid.SelectionChanged -= TechnoFrameColorsDataGrid_SelectionChanged;
            InsetTypesDataGrid.SelectionChanged -= InsetTypesDataGrid_SelectionChanged;
            InsetColorsDataGrid.SelectionChanged -= InsetColorsDataGrid_SelectionChanged;
            TechnoInsetTypesDataGrid.SelectionChanged -= TechnoInsetTypesDataGrid_SelectionChanged;
            TechnoInsetColorsDataGrid.SelectionChanged -= TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged -= PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged -= FrontsHeightDataGrid_SelectionChanged;

            FrontsDataGrid.DataSource = _frontsCatalog.FrontsBindingSource;
            FrameColorsDataGrid.DataSource = _frontsCatalog.FrameColorsBindingSource;
            TechnoFrameColorsDataGrid.DataSource = _frontsCatalog.TechnoFrameColorsBindingSource;
            PatinaDataGrid.DataSource = _frontsCatalog.PatinaBindingSource;
            InsetTypesDataGrid.DataSource = _frontsCatalog.InsetTypesBindingSource;
            InsetColorsDataGrid.DataSource = _frontsCatalog.InsetColorsBindingSource;
            TechnoInsetTypesDataGrid.DataSource = _frontsCatalog.TechnoInsetTypesBindingSource;
            TechnoInsetColorsDataGrid.DataSource = _frontsCatalog.TechnoInsetColorsBindingSource;
            FrontsHeightDataGrid.DataSource = _frontsCatalog.HeightBindingSource;
            FrontsWidthDataGrid.DataSource = _frontsCatalog.WidthBindingSource;

            _frontsCatalog.FilterFronts();
            GetFrameColors();
            GetTechnoFrameColors();
            GetInsetTypes();
            GetInsetColors();
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetPatina();
            GetFrontHeight();
            GetFrontWidth();

            FrontsDataGrid.SelectionChanged += FrontsDataGrid_SelectionChanged;
            FrameColorsDataGrid.SelectionChanged += FrameColorsDataGrid_SelectionChanged;
            TechnoFrameColorsDataGrid.SelectionChanged += TechnoFrameColorsDataGrid_SelectionChanged;
            InsetTypesDataGrid.SelectionChanged += InsetTypesDataGrid_SelectionChanged;
            InsetColorsDataGrid.SelectionChanged += InsetColorsDataGrid_SelectionChanged;
            TechnoInsetTypesDataGrid.SelectionChanged += TechnoInsetTypesDataGrid_SelectionChanged;
            TechnoInsetColorsDataGrid.SelectionChanged += TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged += PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged += FrontsHeightDataGrid_SelectionChanged;
        }

        private void FrontsGridsSettings()
        {
            FrontsDataGrid.Columns["FrontID"].Visible = false;
            FrameColorsDataGrid.Columns["ColorID"].Visible = false;
            PatinaDataGrid.Columns["PatinaID"].Visible = false;
            InsetTypesDataGrid.Columns["InsetTypeID"].Visible = false;
            InsetTypesDataGrid.Columns["GroupID"].Visible = false;
            InsetColorsDataGrid.Columns["table_id"].Visible = false;
            InsetColorsDataGrid.Columns["InsetColorID"].Visible = false;
            InsetColorsDataGrid.Columns["GroupID"].Visible = false;
            TechnoFrameColorsDataGrid.Columns["ColorID"].Visible = false;
            TechnoInsetTypesDataGrid.Columns["InsetTypeID"].Visible = false;
            TechnoInsetTypesDataGrid.Columns["GroupID"].Visible = false;
            TechnoInsetColorsDataGrid.Columns["InsetColorID"].Visible = false;
            TechnoInsetColorsDataGrid.Columns["GroupID"].Visible = false;

            FrontsDataGrid.Columns["FrontName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            FrameColorsDataGrid.Columns["ColorName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            PatinaDataGrid.Columns["PatinaName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            InsetTypesDataGrid.Columns["InsetTypeName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            InsetColorsDataGrid.Columns["InsetColorName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            TechnoFrameColorsDataGrid.Columns["ColorName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            TechnoInsetTypesDataGrid.Columns["InsetTypeName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            TechnoInsetColorsDataGrid.Columns["InsetColorName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        public void GetFrameColors()
        {
            if (_frontsCatalog == null)
                return;
            string frontName = "";
            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();

            int ClientID = -1;
            if (cbClients.Checked && FilterClientsDataGrid.SelectedRows.Count > 0)
                ClientID = Convert.ToInt32(FilterClientsDataGrid.SelectedRows[0].Cells["ClientID"].Value);

            if (cbClients.Checked)
                _frontsCatalog.FilterCatalogFrameColors(ClientID, frontName);
            else
                _frontsCatalog.FilterCatalogFrameColors(frontName);
        }

        public void GetTechnoFrameColors()
        {
            if (_frontsCatalog == null)
                return;
            string frontName = "";
            int colorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);

            _frontsCatalog.FilterCatalogTechnoFrameColors(frontName, colorId);
        }

        public void GetPatina()
        {
            if (_frontsCatalog == null)
                return;
            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            _frontsCatalog.FilterCatalogPatina(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
        }

        public void GetInsetTypes()
        {
            if (_frontsCatalog == null)
                return;
            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);

            _frontsCatalog.FilterCatalogInsetTypes(frontName, colorId, technoColorId);
        }

        public void GetInsetColors()
        {
            if (_frontsCatalog == null)
                return;
            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);

            _frontsCatalog.FilterCatalogInsetColors(frontName, colorId, technoColorId, insetTypeId);
        }

        public void GetTechnoInsetTypes()
        {
            if (_frontsCatalog == null)
                return;
            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            _frontsCatalog.FilterCatalogTechnoInsetTypes(frontName, colorId, technoColorId, insetTypeId, insetColorId);

        }

        public void GetTechnoInsetColors()
        {
            if (_frontsCatalog == null)
                return;
            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);

            _frontsCatalog.FilterCatalogTechnoInsetColors(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId);

        }

        private Bitmap _layer;
        private Bitmap _backgr;
        private Bitmap _backgr2;
        private Bitmap _newBitmap;
        private Bitmap _newBitmap2;

        public void GetFrontHeight()
        {
            if (_frontsCatalog == null)
                return;
            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            _frontsCatalog.FilterCatalogHeight(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);

            pcbxFrontImage.Image = null;
            if (_needSplash)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                int configId = -1;
                if (cbFrontImage.Checked)
                {
                    configId = _frontsCatalog.GetFrontAttachments(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);
                    if (configId != -1)
                        pcbxFrontImage.Image = _frontsCatalog.GetFrontConfigImage(configId);

                    if (pcbxFrontImage.Image == null)
                        pcbxFrontImage.Cursor = Cursors.Default;
                    else
                        pcbxFrontImage.Cursor = Cursors.Hand;
                }
                pcbxFrontTechStore.Image = null;
                if (cbFrontTechStore.Checked)
                {
                    if (frontName.Length > 0)
                    {
                        int frontId = _frontsCatalog.GetFrontID(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);
                        if (frontId != -1)
                            pcbxFrontTechStore.Image = _frontsCatalog.GetTechStoreImage(frontId);
                    }

                    if (pcbxFrontTechStore.Image == null)
                        pcbxFrontTechStore.Cursor = Cursors.Default;
                    else
                        pcbxFrontTechStore.Cursor = Cursors.Hand;
                }
                pcbxVisualConfig.Image = null;
                if (cbVisualConfig.Checked)
                {
                    if (frontName.Length > 0)
                    {
                        _layer = null;
                        _backgr = null;
                        _backgr2 = null;
                        _newBitmap = null;
                        _newBitmap2 = null;
                        configId = _frontsCatalog.GetInsetTypesConfig(frontName, colorId, technoColorId);
                        Image img = _frontsCatalog.GetInsetTypeImage(configId);
                        if (configId != -1 && img != null)
                            _layer = new Bitmap(img);
                        configId = _frontsCatalog.GetInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                        img = _frontsCatalog.GetInsetColorImage(configId);
                        if (configId != -1 && img != null)
                            _backgr = new Bitmap(img);

                        configId = _frontsCatalog.GetTechnoInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                        img = _frontsCatalog.GetTechnoInsetColorImage(configId);
                        if (configId != -1 && img != null)
                            _backgr2 = new Bitmap(img);

                        if (_layer != null)
                        {
                            if (_backgr != null)
                            {
                                if (_backgr2 != null)
                                {
                                    int width = _layer.Width;
                                    int height = _layer.Height;
                                    int width1 = _backgr.Width;
                                    int height1 = _backgr.Height;
                                    int width2 = _backgr2.Width;
                                    int height2 = _backgr2.Height;
                                    _newBitmap = new Bitmap(width, height);
                                    _newBitmap2 = new Bitmap(width1, height1);
                                    using (var canvas = Graphics.FromImage(_newBitmap2))
                                    {
                                        if (height2 >= height1)
                                        {
                                            canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                            canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                            canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                            canvas.Save();
                                        }
                                        else
                                        {
                                            canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                            canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr2.Width, _backgr2.Height), GraphicsUnit.Pixel);
                                            canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                            canvas.Save();
                                        }

                                        int width3 = _newBitmap2.Width;
                                        int height3 = _newBitmap2.Height;
                                        using (var canvas1 = Graphics.FromImage(_newBitmap))
                                        {
                                            if (height3 >= height)
                                            {
                                                canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                                canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                                canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                                canvas1.Save();
                                            }
                                            else
                                            {
                                                canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                                canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _newBitmap2.Width, _newBitmap2.Height), GraphicsUnit.Pixel);
                                                canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                                canvas1.Save();
                                            }
                                        }
                                    }

                                    try
                                    {
                                        pcbxVisualConfig.Image = _newBitmap;
                                    }
                                    catch (Exception ex) { MessageBox.Show(ex.Message); }
                                }
                                else
                                {
                                    int width = _layer.Width;
                                    int height = _layer.Height;
                                    int width1 = _backgr.Width;
                                    int height1 = _backgr.Height;
                                    _newBitmap = new Bitmap(width, height);
                                    using (var canvas = Graphics.FromImage(_newBitmap))
                                    {
                                        if (height1 >= height)
                                        {
                                            canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                            canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                            canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                            canvas.Save();
                                        }
                                        else
                                        {
                                            canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                            canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                            canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                            canvas.Save();
                                        }
                                    }

                                    try
                                    {
                                        pcbxVisualConfig.Image = _newBitmap;
                                    }
                                    catch (Exception ex) { MessageBox.Show(ex.Message); }
                                }
                            }
                            else
                                pcbxVisualConfig.Image = _layer;
                        }
                        else
                        {
                            if (_backgr != null)
                                pcbxVisualConfig.Image = _backgr;
                        }
                        if (pcbxVisualConfig.Image == null)
                            pcbxVisualConfig.Cursor = Cursors.Default;
                        else
                            pcbxVisualConfig.Cursor = Cursors.Hand;
                    }

                }
                //cbFrontToSite.Checked = IsConfigImageToSite;
                //BarcodeLabel.ForeColor = Color.FromArgb(82, 169, 24);
                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

            }
            else
            {
                int configId = -1;

                if (cbFrontImage.Checked)
                {
                    configId = _frontsCatalog.GetFrontAttachments(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);
                    if (configId != -1)
                        pcbxFrontImage.Image = _frontsCatalog.GetFrontConfigImage(configId);

                    if (pcbxFrontImage.Image == null)
                        pcbxFrontImage.Cursor = Cursors.Default;
                    else
                        pcbxFrontImage.Cursor = Cursors.Hand;
                }
                pcbxFrontTechStore.Image = null;
                if (cbFrontTechStore.Checked)
                {
                    if (frontName.Length > 0)
                    {
                        int frontId = _frontsCatalog.GetFrontID(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);
                        if (frontId != -1)
                            pcbxFrontTechStore.Image = _frontsCatalog.GetTechStoreImage(frontId);
                    }

                    if (pcbxFrontTechStore.Image == null)
                        pcbxFrontTechStore.Cursor = Cursors.Default;
                    else
                        pcbxFrontTechStore.Cursor = Cursors.Hand;
                }
                pcbxVisualConfig.Image = null;
                if (cbVisualConfig.Checked)
                {
                    if (frontName.Length > 0)
                    {
                        _layer = null;
                        _backgr = null;
                        _backgr2 = null;
                        _newBitmap = null;
                        _newBitmap2 = null;
                        configId = _frontsCatalog.GetInsetTypesConfig(frontName, colorId, technoColorId);
                        Image img = _frontsCatalog.GetInsetTypeImage(configId);
                        if (configId != -1 && img != null)
                            _layer = new Bitmap(img);
                        configId = _frontsCatalog.GetInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                        img = _frontsCatalog.GetInsetColorImage(configId);
                        if (configId != -1 && img != null)
                            _backgr = new Bitmap(img);

                        configId = _frontsCatalog.GetTechnoInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                        img = _frontsCatalog.GetTechnoInsetColorImage(configId);
                        if (configId != -1 && img != null)
                            _backgr2 = new Bitmap(img);

                        if (_layer != null)
                        {
                            if (_backgr != null)
                            {
                                if (_backgr2 != null)
                                {
                                    int width = _layer.Width;
                                    int height = _layer.Height;
                                    int width1 = _backgr.Width;
                                    int height1 = _backgr.Height;
                                    int width2 = _backgr2.Width;
                                    int height2 = _backgr2.Height;
                                    _newBitmap = new Bitmap(width, height);
                                    _newBitmap2 = new Bitmap(width1, height1);
                                    using (var canvas = Graphics.FromImage(_newBitmap2))
                                    {
                                        if (height2 >= height1)
                                        {
                                            canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                            canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                            canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                            canvas.Save();
                                        }
                                        else
                                        {
                                            canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                            canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr2.Width, _backgr2.Height), GraphicsUnit.Pixel);
                                            canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                            canvas.Save();
                                        }

                                        int width3 = _newBitmap2.Width;
                                        int height3 = _newBitmap2.Height;
                                        using (var canvas1 = Graphics.FromImage(_newBitmap))
                                        {
                                            if (height3 >= height)
                                            {
                                                canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                                canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                                canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                                canvas1.Save();
                                            }
                                            else
                                            {
                                                canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                                canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _newBitmap2.Width, _newBitmap2.Height), GraphicsUnit.Pixel);
                                                canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                                canvas1.Save();
                                            }
                                        }
                                    }

                                    try
                                    {
                                        pcbxVisualConfig.Image = _newBitmap;
                                    }
                                    catch (Exception ex) { MessageBox.Show(ex.Message); }
                                }
                                else
                                {
                                    int width = _layer.Width;
                                    int height = _layer.Height;
                                    int width1 = _backgr.Width;
                                    int height1 = _backgr.Height;
                                    _newBitmap = new Bitmap(width, height);
                                    using (var canvas = Graphics.FromImage(_newBitmap))
                                    {
                                        if (height1 >= height)
                                        {
                                            canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                            canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                            canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                            canvas.Save();
                                        }
                                        else
                                        {
                                            canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                            canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                            canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                            canvas.Save();
                                        }
                                    }

                                    try
                                    {
                                        pcbxVisualConfig.Image = _newBitmap;
                                    }
                                    catch (Exception ex) { MessageBox.Show(ex.Message); }
                                }
                            }
                            else
                                pcbxVisualConfig.Image = _layer;
                        }
                        else
                        {
                            if (_backgr != null)
                                pcbxVisualConfig.Image = _backgr;
                        }
                    }
                    if (pcbxVisualConfig.Image == null)
                        pcbxVisualConfig.Cursor = Cursors.Default;
                    else
                        pcbxVisualConfig.Cursor = Cursors.Hand;
                }
                //cbFrontToSite.Checked = IsConfigImageToSite;
            }
        }

        public void GetFrontWidth()
        {
            if (_frontsCatalog == null)
                return;
            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;
            int height = 0;
            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (FrontsHeightDataGrid.SelectedRows.Count > 0 && FrontsHeightDataGrid.SelectedRows[0].Cells["Height"].Value != DBNull.Value)
                height = Convert.ToInt32(FrontsHeightDataGrid.SelectedRows[0].Cells["Height"].Value);

            _frontsCatalog.FilterCatalogWidth(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId, height);
        }

        #endregion

        #region DECOR FUNCTIONS
        public void DecorCatalogBinding()
        {
            ProductsDataGrid.SelectionChanged -= ProductsDataGrid_SelectionChanged;
            DecorDataGrid.SelectionChanged -= DecorDataGrid_SelectionChanged;
            ColorsDataGrid.SelectionChanged -= ColorsDataGrid_SelectionChanged;
            DecorPatinaDataGrid.SelectionChanged -= DecorPatinaDataGrid_SelectionChanged;
            DecorInsetTypesDataGrid.SelectionChanged -= DecorInsetTypesDataGrid_SelectionChanged;
            DecorInsetColorsDataGrid.SelectionChanged -= DecorInsetColorsDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged -= HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged -= WidthDataGrid_SelectionChanged;
            LengthDataGrid.SelectionChanged -= LengthDataGrid_SelectionChanged;

            ProductsDataGrid.DataSource = _decorCatalog.DecorProductsBindingSource;
            DecorDataGrid.DataSource = _decorCatalog.DecorItemBindingSource;
            ColorsDataGrid.DataSource = _decorCatalog.ItemColorsBindingSource;
            DecorPatinaDataGrid.DataSource = _decorCatalog.ItemPatinaBindingSource;
            DecorInsetTypesDataGrid.DataSource = _decorCatalog.ItemInsetTypesBindingSource;
            DecorInsetColorsDataGrid.DataSource = _decorCatalog.ItemInsetColorsBindingSource;
            LengthDataGrid.DataSource = _decorCatalog.LengthBindingSource;
            HeightDataGrid.DataSource = _decorCatalog.HeightBindingSource;
            WidthDataGrid.DataSource = _decorCatalog.WidthBindingSource;

            GetDecorItems();
            GetDecorColors();
            GetDecorPatina();
            GetDecorInsetTypes();
            GetDecorInsetColors();
            GetDecorLength();
            GetDecorHeight();
            GetDecorWidth();
            ProductsDataGrid.SelectionChanged += ProductsDataGrid_SelectionChanged;
            DecorDataGrid.SelectionChanged += DecorDataGrid_SelectionChanged;
            ColorsDataGrid.SelectionChanged += ColorsDataGrid_SelectionChanged;
            DecorPatinaDataGrid.SelectionChanged += DecorPatinaDataGrid_SelectionChanged;
            DecorInsetTypesDataGrid.SelectionChanged += DecorInsetTypesDataGrid_SelectionChanged;
            DecorInsetColorsDataGrid.SelectionChanged += DecorInsetColorsDataGrid_SelectionChanged;
            LengthDataGrid.SelectionChanged += LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged += HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged += WidthDataGrid_SelectionChanged;
        }

        private void DecorGridsSettings()
        {
            ProductsDataGrid.Columns["ProductID"].Visible = false;
            ProductsDataGrid.Columns["MeasureID"].Visible = false;
            ProductsDataGrid.Columns["ReportParam"].Visible = false;

            //DecorDataGrid.Columns["ProductID"].Visible = false;
            //DecorDataGrid.Columns["DecorID"].Visible = false;

            ColorsDataGrid.Columns["ColorID"].Visible = false;
            //ColorsDataGrid.Columns["GroupID"].Visible = false;

            DecorPatinaDataGrid.Columns["PatinaID"].Visible = false;
            DecorInsetTypesDataGrid.Columns["InsetTypeID"].Visible = false;
            DecorInsetColorsDataGrid.Columns["InsetColorID"].Visible = false;
            DecorInsetTypesDataGrid.Columns["GroupID"].Visible = false;
            DecorInsetColorsDataGrid.Columns["GroupID"].Visible = false;

            ProductsDataGrid.Columns["ProductName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            DecorDataGrid.Columns["Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            ColorsDataGrid.Columns["ColorName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            DecorPatinaDataGrid.Columns["PatinaName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            DecorInsetTypesDataGrid.Columns["InsetTypeName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            DecorInsetColorsDataGrid.Columns["InsetColorName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        public void GetDecorItems()
        {
            if (_decorCatalog != null)
                if (_decorCatalog.DecorProductsBindingSource.Count > 0)
                {
                    if (((DataRowView)_decorCatalog.DecorProductsBindingSource.Current)["ProductID"] != DBNull.Value)
                    {
                        _decorCatalog.FilterItems(Convert.ToInt32(((DataRowView)_decorCatalog.DecorProductsBindingSource.Current)["ProductID"]));
                        _decorCatalog.DecorItemBindingSource.MoveFirst();
                    }
                }
        }

        public void GetDecorColors()
        {
            if (_decorCatalog.DecorItemBindingSource.Current == null)
                return;
            if (_decorCatalog.DecorItemBindingSource.Count > 0)
            {
                _decorCatalog.FilterColors(((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString());
            }
        }

        public void GetDecorPatina()
        {
            if (_decorCatalog.DecorItemBindingSource.Current == null)
                return;
            if (_decorCatalog.ItemColorsBindingSource.Count > 0)
            {
                _decorCatalog.FilterPatina(((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString(),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]));
            }
        }

        public void GetDecorInsetTypes()
        {
            if (_decorCatalog.DecorItemBindingSource.Current == null)
                return;
            if (_decorCatalog.ItemPatinaBindingSource.Count > 0)
            {
                _decorCatalog.FilterInsetType(((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString(),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaID"]));
            }
        }

        public void GetDecorInsetColors()
        {
            if (_decorCatalog.DecorItemBindingSource.Current == null)
                return;
            if (_decorCatalog.ItemInsetTypesBindingSource.Count > 0)
            {
                _decorCatalog.FilterInsetColor(((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString(),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current)["InsetTypeID"]));
            }
        }

        public void GetDecorLength()
        {
            if (_decorCatalog.DecorItemBindingSource.Current == null)
                return;
            if (_decorCatalog.ItemInsetColorsBindingSource.Count > 0)
            {

                _decorCatalog.FilterLength(((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString(),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current)["InsetTypeID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current)["InsetColorID"]));
            }
        }

        public void GetDecorHeight()
        {
            if (_decorCatalog.DecorProductsBindingSource.Current == null || _decorCatalog.DecorItemBindingSource.Current == null)
                return;
            if (_decorCatalog.LengthBindingSource.Count > 0)
            {
                int length = Convert.ToInt32(((DataRowView)_decorCatalog.LengthBindingSource.Current)["Length"]);
                _decorCatalog.FilterHeight(((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString(),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current)["InsetTypeID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current)["InsetColorID"]), length);
            }

            int productId = -1;
            string decorName = string.Empty;
            int colorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;

            if ((DataRowView)_decorCatalog.DecorProductsBindingSource.Current != null)
                productId = Convert.ToInt32(((DataRowView)_decorCatalog.DecorProductsBindingSource.Current)["ProductID"]);
            if ((DataRowView)_decorCatalog.DecorItemBindingSource.Current != null)
                decorName = ((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString();
            if ((DataRowView)_decorCatalog.ItemColorsBindingSource.Current != null)
                colorId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]);
            if ((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current != null)
                patinaId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaID"]);
            if ((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current != null)
                insetTypeId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current)["InsetTypeID"]);
            if ((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current != null)
                insetColorId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current)["InsetColorID"]);

            if (cbDecorTechStore.Checked)
            {
                //Thread T = new Thread(delegate() { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                //T.Start();

                //while (!SplashWindow.bSmallCreated) ;

                //int DecorID = DecorCatalog.GetDecorID(ProductID, DecorName, ColorID, PatinaID);
                //if (DecorID != -1)
                //    pcbxDecorTechStore.Image = DecorCatalog.GetTechStoreImage(DecorID);

                //while (SplashWindow.bSmallCreated)
                //    SmallWaitForm.CloseS = true;

                int decorId = _decorCatalog.GetDecorID(productId, decorName, colorId, patinaId, insetTypeId, insetColorId);
                if (decorId != -1)
                    pcbxDecorTechStore.Image = _decorCatalog.GetTechStoreImage(decorId);

                if (pcbxDecorTechStore.Image == null)
                    pcbxDecorTechStore.Cursor = Cursors.Default;
                else
                    pcbxDecorTechStore.Cursor = Cursors.Hand;
            }
            pcbxDecorImage.Image = null;
            if (cbDecorImage.Checked)
            {
                //Thread T = new Thread(delegate() { SplashWindow.CreateSmallSplash(ref TopForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                //T.Start();

                //while (!SplashWindow.bSmallCreated) ;

                //int ConfigID = DecorCatalog.SaveDecorAttachments(ProductID, DecorName, ColorID, PatinaID);
                //if (ConfigID != -1)
                //    pcbxDecorImage.Image = DecorCatalog.GetDecorConfigImage(ConfigID);

                //while (SplashWindow.bSmallCreated)
                //    SmallWaitForm.CloseS = true;

                int configId = _decorCatalog.GetDecorAttachments(productId, decorName, colorId, patinaId, insetTypeId, insetColorId);
                if (configId != -1)
                    pcbxDecorImage.Image = _decorCatalog.GetDecorConfigImage(configId, productId);

                if (pcbxDecorImage.Image == null)
                    pcbxDecorImage.Cursor = Cursors.Default;
                else
                    pcbxDecorImage.Cursor = Cursors.Hand;
            }
        }

        public void GetDecorWidth()
        {
            if (_decorCatalog.DecorItemBindingSource.Current == null)
                return;
            if (_decorCatalog.HeightBindingSource.Count > 0)
            {
                int length = Convert.ToInt32(((DataRowView)_decorCatalog.LengthBindingSource.Current)["Length"]);
                int height = Convert.ToInt32(((DataRowView)_decorCatalog.HeightBindingSource.Current)["Height"]);

                _decorCatalog.FilterWidth(((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString(),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current)["InsetTypeID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current)["InsetColorID"]), length, height);
            }

        }

        #endregion


        private void ProductsDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DecorDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void ColorsDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void DecorPatinaDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void LengthDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void HeightDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }


        private void kryptonCheckSet1_CheckedButtonChanged(object sender, EventArgs e)
        {
            if (_frontsCatalog == null || _decorCatalog == null)
                return;

            if (kryptonCheckSet1.CheckedButton.Name == "DecorCheckButton")
            {
                ProductsDataGrid.SelectionChanged -= ProductsDataGrid_SelectionChanged;
                DecorDataGrid.SelectionChanged -= DecorDataGrid_SelectionChanged;
                ColorsDataGrid.SelectionChanged -= ColorsDataGrid_SelectionChanged;
                DecorPatinaDataGrid.SelectionChanged -= DecorPatinaDataGrid_SelectionChanged;
                DecorInsetTypesDataGrid.SelectionChanged -= DecorInsetTypesDataGrid_SelectionChanged;
                DecorInsetColorsDataGrid.SelectionChanged -= DecorInsetColorsDataGrid_SelectionChanged;
                LengthDataGrid.SelectionChanged -= LengthDataGrid_SelectionChanged;
                HeightDataGrid.SelectionChanged -= HeightDataGrid_SelectionChanged;
                WidthDataGrid.SelectionChanged -= WidthDataGrid_SelectionChanged;
                _decorCatalog.FilterProducts();
                GetDecorItems();
                GetDecorColors();
                GetDecorPatina();
                GetDecorInsetTypes();
                GetDecorInsetColors();
                GetDecorLength();
                GetDecorHeight();
                GetDecorWidth();
                ProductsDataGrid.SelectionChanged += ProductsDataGrid_SelectionChanged;
                DecorDataGrid.SelectionChanged += DecorDataGrid_SelectionChanged;
                ColorsDataGrid.SelectionChanged += ColorsDataGrid_SelectionChanged;
                DecorPatinaDataGrid.SelectionChanged += DecorPatinaDataGrid_SelectionChanged;
                DecorInsetTypesDataGrid.SelectionChanged += DecorInsetTypesDataGrid_SelectionChanged;
                DecorInsetColorsDataGrid.SelectionChanged += DecorInsetColorsDataGrid_SelectionChanged;
                LengthDataGrid.SelectionChanged += LengthDataGrid_SelectionChanged;
                HeightDataGrid.SelectionChanged += HeightDataGrid_SelectionChanged;
                WidthDataGrid.SelectionChanged += WidthDataGrid_SelectionChanged;
                DecorCatalogPanel.BringToFront();
            }
            else
            {
                FrontsDataGrid.SelectionChanged -= FrontsDataGrid_SelectionChanged;
                FrameColorsDataGrid.SelectionChanged -= FrameColorsDataGrid_SelectionChanged;
                TechnoFrameColorsDataGrid.SelectionChanged -= TechnoFrameColorsDataGrid_SelectionChanged;
                InsetTypesDataGrid.SelectionChanged -= InsetTypesDataGrid_SelectionChanged;
                InsetColorsDataGrid.SelectionChanged -= InsetColorsDataGrid_SelectionChanged;
                TechnoInsetTypesDataGrid.SelectionChanged -= TechnoInsetTypesDataGrid_SelectionChanged;
                TechnoInsetColorsDataGrid.SelectionChanged -= TechnoInsetColorsDataGrid_SelectionChanged;
                PatinaDataGrid.SelectionChanged -= PatinaDataGrid_SelectionChanged;
                FrontsHeightDataGrid.SelectionChanged -= FrontsHeightDataGrid_SelectionChanged;

                int ClientID = -1;

                if (cbClients.Checked && FilterClientsDataGrid.SelectedRows.Count > 0)
                    ClientID = Convert.ToInt32(FilterClientsDataGrid.SelectedRows[0].Cells["ClientID"].Value);

                if (cbClients.Checked)
                    _frontsCatalog.FilterFronts(ClientID);
                else
                    _frontsCatalog.FilterFronts();
                GetFrameColors();
                GetTechnoFrameColors();
                GetInsetTypes();
                GetInsetColors();
                GetTechnoInsetTypes();
                GetTechnoInsetColors();
                GetPatina();
                GetFrontHeight();
                GetFrontWidth();

                FrontsDataGrid.SelectionChanged += FrontsDataGrid_SelectionChanged;
                FrameColorsDataGrid.SelectionChanged += FrameColorsDataGrid_SelectionChanged;
                TechnoFrameColorsDataGrid.SelectionChanged += TechnoFrameColorsDataGrid_SelectionChanged;
                InsetTypesDataGrid.SelectionChanged += InsetTypesDataGrid_SelectionChanged;
                InsetColorsDataGrid.SelectionChanged += InsetColorsDataGrid_SelectionChanged;
                TechnoInsetTypesDataGrid.SelectionChanged += TechnoInsetTypesDataGrid_SelectionChanged;
                TechnoInsetColorsDataGrid.SelectionChanged += TechnoInsetColorsDataGrid_SelectionChanged;
                PatinaDataGrid.SelectionChanged += PatinaDataGrid_SelectionChanged;
                FrontsHeightDataGrid.SelectionChanged += FrontsHeightDataGrid_SelectionChanged;
                FrontsCatalogPanel.BringToFront();
            }
        }

        private void kryptonCheckSet2_CheckedButtonChanged(object sender, EventArgs e)
        {
            if (_frontsCatalog == null || _decorCatalog == null)
                return;

            int ClientID = -1;

            if (kryptonCheckSet2.CheckedButton.Name == "ProfilCheckButton")
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                _frontsCatalog.FilterCatalog(1);
                _decorCatalog.FilterCatalog(1);

                FrontsDataGrid.SelectionChanged -= FrontsDataGrid_SelectionChanged;
                FrameColorsDataGrid.SelectionChanged -= FrameColorsDataGrid_SelectionChanged;
                TechnoFrameColorsDataGrid.SelectionChanged -= TechnoFrameColorsDataGrid_SelectionChanged;
                InsetTypesDataGrid.SelectionChanged -= InsetTypesDataGrid_SelectionChanged;
                InsetColorsDataGrid.SelectionChanged -= InsetColorsDataGrid_SelectionChanged;
                TechnoInsetTypesDataGrid.SelectionChanged -= TechnoInsetTypesDataGrid_SelectionChanged;
                TechnoInsetColorsDataGrid.SelectionChanged -= TechnoInsetColorsDataGrid_SelectionChanged;
                PatinaDataGrid.SelectionChanged -= PatinaDataGrid_SelectionChanged;
                FrontsHeightDataGrid.SelectionChanged -= FrontsHeightDataGrid_SelectionChanged;

                if (cbClients.Checked && FilterClientsDataGrid.SelectedRows.Count > 0)
                    ClientID = Convert.ToInt32(FilterClientsDataGrid.SelectedRows[0].Cells["ClientID"].Value);

                if (cbClients.Checked)
                    _frontsCatalog.FilterFronts(ClientID);
                else
                    _frontsCatalog.FilterFronts();

                GetFrameColors();
                GetTechnoFrameColors();
                GetInsetTypes();
                GetInsetColors();
                GetTechnoInsetTypes();
                GetTechnoInsetColors();
                GetPatina();
                GetFrontHeight();
                GetFrontWidth();
                FrontsDataGrid.SelectionChanged += FrontsDataGrid_SelectionChanged;
                FrameColorsDataGrid.SelectionChanged += FrameColorsDataGrid_SelectionChanged;
                TechnoFrameColorsDataGrid.SelectionChanged += TechnoFrameColorsDataGrid_SelectionChanged;
                InsetTypesDataGrid.SelectionChanged += InsetTypesDataGrid_SelectionChanged;
                InsetColorsDataGrid.SelectionChanged += InsetColorsDataGrid_SelectionChanged;
                TechnoInsetTypesDataGrid.SelectionChanged += TechnoInsetTypesDataGrid_SelectionChanged;
                TechnoInsetColorsDataGrid.SelectionChanged += TechnoInsetColorsDataGrid_SelectionChanged;
                PatinaDataGrid.SelectionChanged += PatinaDataGrid_SelectionChanged;
                FrontsHeightDataGrid.SelectionChanged += FrontsHeightDataGrid_SelectionChanged;

                ProductsDataGrid.SelectionChanged -= ProductsDataGrid_SelectionChanged;
                DecorDataGrid.SelectionChanged -= DecorDataGrid_SelectionChanged;
                ColorsDataGrid.SelectionChanged -= ColorsDataGrid_SelectionChanged;
                DecorPatinaDataGrid.SelectionChanged -= DecorPatinaDataGrid_SelectionChanged;
                DecorInsetTypesDataGrid.SelectionChanged -= DecorInsetTypesDataGrid_SelectionChanged;
                DecorInsetColorsDataGrid.SelectionChanged -= DecorInsetColorsDataGrid_SelectionChanged;
                HeightDataGrid.SelectionChanged -= HeightDataGrid_SelectionChanged;
                WidthDataGrid.SelectionChanged -= WidthDataGrid_SelectionChanged;
                LengthDataGrid.SelectionChanged -= LengthDataGrid_SelectionChanged;
                _decorCatalog.FilterProducts();
                GetDecorItems();
                GetDecorColors();
                GetDecorPatina();
                GetDecorInsetTypes();
                GetDecorInsetColors();
                GetDecorLength();
                GetDecorHeight();
                GetDecorWidth();
                ProductsDataGrid.SelectionChanged += ProductsDataGrid_SelectionChanged;
                DecorDataGrid.SelectionChanged += DecorDataGrid_SelectionChanged;
                ColorsDataGrid.SelectionChanged += ColorsDataGrid_SelectionChanged;
                DecorPatinaDataGrid.SelectionChanged += DecorPatinaDataGrid_SelectionChanged;
                DecorInsetTypesDataGrid.SelectionChanged += DecorInsetTypesDataGrid_SelectionChanged;
                DecorInsetColorsDataGrid.SelectionChanged += DecorInsetColorsDataGrid_SelectionChanged;
                HeightDataGrid.SelectionChanged += HeightDataGrid_SelectionChanged;
                WidthDataGrid.SelectionChanged += WidthDataGrid_SelectionChanged;
                LengthDataGrid.SelectionChanged += LengthDataGrid_SelectionChanged;

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                _frontsCatalog.FilterCatalog(2);
                _decorCatalog.FilterCatalog(2);

                FrontsDataGrid.SelectionChanged -= FrontsDataGrid_SelectionChanged;
                FrameColorsDataGrid.SelectionChanged -= FrameColorsDataGrid_SelectionChanged;
                TechnoFrameColorsDataGrid.SelectionChanged -= TechnoFrameColorsDataGrid_SelectionChanged;
                InsetTypesDataGrid.SelectionChanged -= InsetTypesDataGrid_SelectionChanged;
                InsetColorsDataGrid.SelectionChanged -= InsetColorsDataGrid_SelectionChanged;
                TechnoInsetTypesDataGrid.SelectionChanged -= TechnoInsetTypesDataGrid_SelectionChanged;
                TechnoInsetColorsDataGrid.SelectionChanged -= TechnoInsetColorsDataGrid_SelectionChanged;
                PatinaDataGrid.SelectionChanged -= PatinaDataGrid_SelectionChanged;
                FrontsHeightDataGrid.SelectionChanged -= FrontsHeightDataGrid_SelectionChanged;

                if (cbClients.Checked && FilterClientsDataGrid.SelectedRows.Count > 0)
                    ClientID = Convert.ToInt32(FilterClientsDataGrid.SelectedRows[0].Cells["ClientID"].Value);

                if (cbClients.Checked)
                    _frontsCatalog.FilterFronts(ClientID);
                else
                    _frontsCatalog.FilterFronts();

                GetFrameColors();
                GetTechnoFrameColors();
                GetInsetTypes();
                GetInsetColors();
                GetTechnoInsetTypes();
                GetTechnoInsetColors();
                GetPatina();
                GetFrontHeight();
                GetFrontWidth();
                FrontsDataGrid.SelectionChanged += FrontsDataGrid_SelectionChanged;
                FrameColorsDataGrid.SelectionChanged += FrameColorsDataGrid_SelectionChanged;
                TechnoFrameColorsDataGrid.SelectionChanged += TechnoFrameColorsDataGrid_SelectionChanged;
                InsetTypesDataGrid.SelectionChanged += InsetTypesDataGrid_SelectionChanged;
                InsetColorsDataGrid.SelectionChanged += InsetColorsDataGrid_SelectionChanged;
                TechnoInsetTypesDataGrid.SelectionChanged += TechnoInsetTypesDataGrid_SelectionChanged;
                TechnoInsetColorsDataGrid.SelectionChanged += TechnoInsetColorsDataGrid_SelectionChanged;
                PatinaDataGrid.SelectionChanged += PatinaDataGrid_SelectionChanged;
                FrontsHeightDataGrid.SelectionChanged += FrontsHeightDataGrid_SelectionChanged;

                ProductsDataGrid.SelectionChanged -= ProductsDataGrid_SelectionChanged;
                DecorDataGrid.SelectionChanged -= DecorDataGrid_SelectionChanged;
                ColorsDataGrid.SelectionChanged -= ColorsDataGrid_SelectionChanged;
                DecorPatinaDataGrid.SelectionChanged -= DecorPatinaDataGrid_SelectionChanged;
                DecorInsetTypesDataGrid.SelectionChanged -= DecorInsetTypesDataGrid_SelectionChanged;
                DecorInsetColorsDataGrid.SelectionChanged -= DecorInsetColorsDataGrid_SelectionChanged;
                HeightDataGrid.SelectionChanged -= HeightDataGrid_SelectionChanged;
                WidthDataGrid.SelectionChanged -= WidthDataGrid_SelectionChanged;
                LengthDataGrid.SelectionChanged -= LengthDataGrid_SelectionChanged;
                _decorCatalog.FilterProducts();
                GetDecorItems();
                GetDecorColors();
                GetDecorPatina();
                GetDecorInsetTypes();
                GetDecorInsetColors();
                GetDecorLength();
                GetDecorHeight();
                GetDecorWidth();
                ProductsDataGrid.SelectionChanged += ProductsDataGrid_SelectionChanged;
                DecorDataGrid.SelectionChanged += DecorDataGrid_SelectionChanged;
                ColorsDataGrid.SelectionChanged += ColorsDataGrid_SelectionChanged;
                DecorPatinaDataGrid.SelectionChanged += DecorPatinaDataGrid_SelectionChanged;
                DecorInsetTypesDataGrid.SelectionChanged += DecorInsetTypesDataGrid_SelectionChanged;
                DecorInsetColorsDataGrid.SelectionChanged += DecorInsetColorsDataGrid_SelectionChanged;
                HeightDataGrid.SelectionChanged += HeightDataGrid_SelectionChanged;
                WidthDataGrid.SelectionChanged += WidthDataGrid_SelectionChanged;
                LengthDataGrid.SelectionChanged += LengthDataGrid_SelectionChanged;

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == 6 && m.WParam.ToInt32() == 1)
            {
                if (_topForm != null)
                    _topForm.Activate();
            }
        }

        private void AddPictureButton_Click(object sender, EventArgs e)
        {
            if (DecorDataGrid.SelectedRows.Count == 0)
                return;

            PictureEditForm pictureEditForm = new PictureEditForm(ref _decorCatalog, Convert.ToInt32(DecorDataGrid.SelectedRows[0].Cells["DecorID"].Value));

            _topForm = pictureEditForm;

            pictureEditForm.ShowDialog();

            _topForm = null;

            pictureEditForm.Dispose();
        }

        private void PrintPackageContextMenuItem_Click(object sender, EventArgs e)
        {
            bool pressOk = false;
            int labelsCount = 0;
            int positionsCount = 0;
            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            SampleFrontsLabelInfoMenu sampleLabelInfoMenu = new SampleFrontsLabelInfoMenu(this);
            _topForm = sampleLabelInfoMenu;
            sampleLabelInfoMenu.ShowDialog();
            pressOk = sampleLabelInfoMenu.PressOK;
            labelsCount = sampleLabelInfoMenu.LabelsCount;
            positionsCount = sampleLabelInfoMenu.PositionsCount;
            phantomForm.Close();
            phantomForm.Dispose();
            sampleLabelInfoMenu.Dispose();
            _topForm = null;

            if (!pressOk)
                return;

            int frontId = -1;
            int factoryId = -1;
            string front = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].FormattedValue.ToString();
            string frameColor = FrameColorsDataGrid.SelectedRows[0].Cells["ColorName"].FormattedValue.ToString();
            string patina = PatinaDataGrid.SelectedRows[0].Cells["PatinaName"].FormattedValue.ToString();
            string insetType = InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeName"].FormattedValue.ToString();
            string insetColor = InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorName"].FormattedValue.ToString();
            string technoInsetType = TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeName"].FormattedValue.ToString();
            string technoInsetColor = TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorName"].FormattedValue.ToString();

            if (patina != "-")
                frameColor += "/" + patina;
            if (technoInsetType != "-")
                insetType += "/" + technoInsetType;
            if (technoInsetColor != "-")
                insetColor += "/" + technoInsetColor;

            int height = 0;
            int width = 0;
            int frontConfigId = 0;

            if (_frontsCatalog.HeightBindingSource.Count > 0)
                height = Convert.ToInt32(((DataRowView)_frontsCatalog.HeightBindingSource.Current).Row["Height"]);

            if (_frontsCatalog.WidthBindingSource.Count > 0)
                width = Convert.ToInt32(((DataRowView)_frontsCatalog.WidthBindingSource.Current).Row["Width"]);

            if (_frontsCatalog != null)
                if (_frontsCatalog.WidthBindingSource.Count > 0)
                {
                    if (((DataRowView)_frontsCatalog.WidthBindingSource.Current)["Width"] != DBNull.Value)
                    {
                        frontConfigId = _frontsCatalog.GetFrontConfigID(((DataRowView)_frontsCatalog.FrontsBindingSource.Current)["FrontName"].ToString(),
                            Convert.ToInt32(((DataRowView)_frontsCatalog.FrameColorsBindingSource.Current)["ColorID"]),
                            Convert.ToInt32(((DataRowView)_frontsCatalog.TechnoFrameColorsBindingSource.Current)["ColorID"]),
                            Convert.ToInt32(((DataRowView)_frontsCatalog.PatinaBindingSource.Current)["PatinaID"]),
                            Convert.ToInt32(((DataRowView)_frontsCatalog.InsetTypesBindingSource.Current)["InsetTypeID"]),
                            Convert.ToInt32(((DataRowView)_frontsCatalog.InsetColorsBindingSource.Current)["InsetColorID"]),
                            Convert.ToInt32(((DataRowView)_frontsCatalog.TechnoInsetTypesBindingSource.Current)["InsetTypeID"]),
                            Convert.ToInt32(((DataRowView)_frontsCatalog.TechnoInsetColorsBindingSource.Current)["InsetColorID"]),
                            height, width, ref frontId, ref factoryId);
                    }
                }

            DataRow newRow = _frontsDt.NewRow();

            newRow["FrameColor"] = frameColor;
            newRow["InsetType"] = insetType;
            newRow["InsetColor"] = insetColor;
            newRow["Front"] = front;
            newRow["LabelsCount"] = labelsCount;
            newRow["PositionsCount"] = positionsCount;
            newRow["FrontConfigID"] = frontConfigId;
            newRow["FrontID"] = frontId;
            newRow["ColorID"] = Convert.ToInt32(((DataRowView)_frontsCatalog.FrameColorsBindingSource.Current)["ColorID"]);
            newRow["TechnoColorID"] = Convert.ToInt32(((DataRowView)_frontsCatalog.TechnoFrameColorsBindingSource.Current)["ColorID"]);

            if (kryptonCheckSet2.CheckedIndex == 0)
                newRow["FactoryType"] = 2;
            if (kryptonCheckSet2.CheckedIndex == 1)
                newRow["FactoryType"] = 1;
            _frontsDt.Rows.Add(newRow);
        }

        private void KryptonContextMenuItem1_Click(object sender, EventArgs e)
        {
            _sampleLabelManager.ClearOrderData();
            _frontsDt.Clear();
        }

        private void kryptonContextMenuItem2_Click(object sender, EventArgs e)
        {
            _sampleLabelManager.ClearLabelInfo();

            //Проверка
            if (_frontsDt.Rows.Count == 0)
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                    "Очередь для печати пуста",
                    "Ошибка печати");
                return;
            }

            for (int i = 0; i < _frontsDt.Rows.Count; i++)
            {
                int labelsCount = Convert.ToInt32(_frontsDt.Rows[i]["LabelsCount"]);
                int frontConfigId = Convert.ToInt32(_frontsDt.Rows[i]["FrontConfigID"]);
                int frontId = 0;
                int colorId = 0;
                int technoColorId = 0;
                int sampleLabelId = 0;
                if (_frontsDt.Rows[i]["FrontID"] != DBNull.Value)
                    frontId = Convert.ToInt32(_frontsDt.Rows[i]["FrontID"]);
                if (_frontsDt.Rows[i]["ColorID"] != DBNull.Value)
                    colorId = Convert.ToInt32(_frontsDt.Rows[i]["ColorID"]);
                if (_frontsDt.Rows[i]["TechnoColorID"] != DBNull.Value)
                    technoColorId = Convert.ToInt32(_frontsDt.Rows[i]["TechnoColorID"]);
                for (int j = 0; j < labelsCount; j++)
                {
                    SampleInfo labelInfo = new SampleInfo();

                    DataTable dt = _frontsDt.Clone();

                    DataRow destRow = dt.NewRow();
                    foreach (DataColumn column in _frontsDt.Columns)
                    {
                        if (column.ColumnName == "FactoryType")
                            continue;
                        destRow[column.ColumnName] = _frontsDt.Rows[i][column.ColumnName];
                    }
                    labelInfo.ProductType = 0;
                    labelInfo.Impost = string.Empty;
                    if ((frontId == 3630 || frontId == 15003) && (colorId == technoColorId))
                        labelInfo.Impost = " с импостом";
                    sampleLabelId = _sampleLabelManager.SaveSampleLabel(frontConfigId, DateTime.Now, Security.CurrentUserID, 0);
                    labelInfo.BarcodeNumber = _sampleLabelManager.GetBarcodeNumber(17, sampleLabelId);
                    labelInfo.FactoryType = Convert.ToInt32(_frontsDt.Rows[i]["FactoryType"]);
                    labelInfo.DocDateTime = DateTime.Now.ToString("dd.MM.yyyy");
                    dt.Rows.Add(destRow);
                    labelInfo.OrderData = dt;

                    _sampleLabelManager.AddLabelInfo(ref labelInfo);
                }
            }
            PrintDialog.Document = _sampleLabelManager.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                _sampleLabelManager.Print();
            }
        }

        private void FrontsDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (_bCanEdit && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu3.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            _formEvent = EHide;
            AnimateTimer.Enabled = true;
        }

        private void kryptonContextMenuItem3_Click(object sender, EventArgs e)
        {
            string front = string.Empty;
            string frameColor = string.Empty;
            string patina = string.Empty;
            string insetType = string.Empty;
            string insetColor = string.Empty;
            string technoInsetType = string.Empty;
            string technoInsetColor = string.Empty;

            string height = string.Empty;
            string width = string.Empty;

            bool pressOk = false;
            string labelsCount = string.Empty;
            string positionsCount = string.Empty;
            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            SampleFrontsManually sampleLabelInfoMenu = new SampleFrontsManually(this);
            _topForm = sampleLabelInfoMenu;
            sampleLabelInfoMenu.ShowDialog();
            front = sampleLabelInfoMenu.Front;
            frameColor = sampleLabelInfoMenu.FrameColor;
            patina = sampleLabelInfoMenu.Patina;
            insetType = sampleLabelInfoMenu.InsetType;
            insetColor = sampleLabelInfoMenu.InsetColor;
            technoInsetType = sampleLabelInfoMenu.TechnoInsetType;
            technoInsetColor = sampleLabelInfoMenu.TechnoInsetColor;
            height = sampleLabelInfoMenu.LabelsHeight;
            width = sampleLabelInfoMenu.LabelsWidth;
            labelsCount = sampleLabelInfoMenu.LabelsCount;
            positionsCount = sampleLabelInfoMenu.PositionsCount;
            pressOk = sampleLabelInfoMenu.PressOK;

            phantomForm.Close();
            phantomForm.Dispose();
            sampleLabelInfoMenu.Dispose();
            _topForm = null;

            if (!pressOk)
                return;
            if (labelsCount.Length == 0)
                return;
            int frontConfigId = 0;
            DataRow newRow = _frontsDt.NewRow();

            newRow["FrameColor"] = frameColor;
            newRow["InsetType"] = insetType;
            newRow["InsetColor"] = insetColor;
            newRow["Front"] = front;
            newRow["LabelsCount"] = labelsCount;
            newRow["PositionsCount"] = positionsCount;
            newRow["FrontConfigID"] = frontConfigId;
            //NewRow["FrontID"] = Convert.ToInt32(((DataRowView)FrontsCatalog.FrontsBindingSource.Current)["FrontID"]);
            //NewRow["ColorID"] = Convert.ToInt32(((DataRowView)FrontsCatalog.FrameColorsBindingSource.Current)["ColorID"]);
            //NewRow["TechnoColorID"] = Convert.ToInt32(((DataRowView)FrontsCatalog.TechnoFrameColorsBindingSource.Current)["ColorID"]);

            if (kryptonCheckSet2.CheckedIndex == 0)
                newRow["FactoryType"] = 2;
            if (kryptonCheckSet2.CheckedIndex == 1)
                newRow["FactoryType"] = 1;
            _frontsDt.Rows.Add(newRow);
        }

        private void ProductsDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (_bCanEdit && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu2.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
                int productId = -1;
                if (ProductsDataGrid.SelectedRows.Count > 0 && ProductsDataGrid.SelectedRows[0].Cells["ProductID"].Value != DBNull.Value)
                    productId = Convert.ToInt32(ProductsDataGrid.SelectedRows[0].Cells["ProductID"].Value);
                kryptonContextMenuItem21.Enabled = false;
                kryptonContextMenuItem22.Enabled = false;
                kryptonContextMenuItem23.Enabled = false;
                kryptonContextMenuItem24.Enabled = false;
                kryptonContextMenuItem25.Enabled = false;
                kryptonContextMenuItem26.Enabled = false;
                kryptonContextMenuItem28.Enabled = false;
                kryptonContextMenuItem29.Enabled = false;
                if (productId == 46) // патриция 1.x
                {
                    kryptonContextMenuItem21.Enabled = true;
                }
                if (productId == 61) // куб
                {
                    kryptonContextMenuItem22.Enabled = true;
                }
                if (productId == 63) // норманн
                {
                    kryptonContextMenuItem23.Enabled = true;
                }
                if (productId == 73) // патриция 1.0
                {
                    kryptonContextMenuItem24.Enabled = true;
                }
                if (productId == 74) // патриция 2.x
                {
                    kryptonContextMenuItem25.Enabled = true;
                }
                if (productId == 75) // патриция 3.x
                {
                    kryptonContextMenuItem26.Enabled = true;
                }
                if (productId == 80) // ПАТРИЦИЯ-4.0 (С)
                {
                    kryptonContextMenuItem28.Enabled = true;
                }
                if (productId == 82) // КРИСМАР
                {
                    kryptonContextMenuItem29.Enabled = true;
                }
            }
        }

        private void kryptonContextMenuItem8_Click(object sender, EventArgs e)
        {
            int decorConfigId = 0;
            bool pressOk = false;

            string labelsCount = string.Empty;
            string positionsCount = string.Empty;
            string labelsLength = string.Empty;
            string labelsHeight = string.Empty;
            string labelsWidth = string.Empty;
            string product = string.Empty;
            string decor = string.Empty;
            string color = string.Empty;
            string patina = string.Empty;

            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            SampleDecorManually sampleLabelInfo = new SampleDecorManually(this);
            _topForm = sampleLabelInfo;
            sampleLabelInfo.ShowDialog();
            pressOk = sampleLabelInfo.PressOK;

            product = sampleLabelInfo.Product;
            decor = sampleLabelInfo.Decor;
            color = sampleLabelInfo.Color;
            patina = sampleLabelInfo.Patina;
            labelsCount = sampleLabelInfo.LabelsCount;
            labelsLength = sampleLabelInfo.LabelsLength;
            labelsHeight = sampleLabelInfo.LabelsHeight;
            labelsWidth = sampleLabelInfo.LabelsWidth;
            positionsCount = sampleLabelInfo.PositionsCount;

            if (patina.Length > 0 || patina != "-")
                color += " " + patina;

            phantomForm.Close();
            phantomForm.Dispose();
            sampleLabelInfo.Dispose();
            _topForm = null;

            if (!pressOk)
                return;

            DataRow newRow = _decorDt.NewRow();

            newRow["Decor"] = decor;
            newRow["Color"] = color;
            newRow["Product"] = product;
            newRow["Length"] = labelsLength.ToString();
            newRow["Height"] = labelsHeight.ToString();
            newRow["Width"] = labelsWidth.ToString();
            newRow["LabelsCount"] = labelsCount;
            newRow["PositionsCount"] = positionsCount;
            newRow["DecorConfigID"] = decorConfigId;

            if (kryptonCheckSet2.CheckedIndex == 0)
                newRow["FactoryType"] = 2;
            if (kryptonCheckSet2.CheckedIndex == 1)
                newRow["FactoryType"] = 1;
            _decorDt.Rows.Add(newRow);
        }

        private void ColorsDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (_bCanEdit && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu2.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
                int productId = -1;
                if (ProductsDataGrid.SelectedRows.Count > 0 && ProductsDataGrid.SelectedRows[0].Cells["ProductID"].Value != DBNull.Value)
                    productId = Convert.ToInt32(ProductsDataGrid.SelectedRows[0].Cells["ProductID"].Value);
                kryptonContextMenuItem21.Enabled = false;
                kryptonContextMenuItem22.Enabled = false;
                kryptonContextMenuItem23.Enabled = false;
                kryptonContextMenuItem24.Enabled = false;
                kryptonContextMenuItem25.Enabled = false;
                kryptonContextMenuItem26.Enabled = false;
                kryptonContextMenuItem28.Enabled = false;
                kryptonContextMenuItem29.Enabled = false;
                if (productId == 46) // патриция 1.x
                {
                    kryptonContextMenuItem21.Enabled = true;
                }
                if (productId == 61) // куб
                {
                    kryptonContextMenuItem22.Enabled = true;
                }
                if (productId == 63) // норманн
                {
                    kryptonContextMenuItem23.Enabled = true;
                }
                if (productId == 73) // патриция 1.0
                {
                    kryptonContextMenuItem24.Enabled = true;
                }
                if (productId == 74) // патриция 2.x
                {
                    kryptonContextMenuItem25.Enabled = true;
                }
                if (productId == 75) // патриция 3.x
                {
                    kryptonContextMenuItem26.Enabled = true;
                }
                if (productId == 80) // ПАТРИЦИЯ-4.0 (С)
                {
                    kryptonContextMenuItem28.Enabled = true;
                }
                if (productId == 82) // КРИСМАР
                {
                    kryptonContextMenuItem29.Enabled = true;
                }
            }
        }

        private void DecorDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (_bCanEdit && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu2.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));

                int productId = -1;
                if (ProductsDataGrid.SelectedRows.Count > 0 && ProductsDataGrid.SelectedRows[0].Cells["ProductID"].Value != DBNull.Value)
                    productId = Convert.ToInt32(ProductsDataGrid.SelectedRows[0].Cells["ProductID"].Value);
                kryptonContextMenuItem21.Enabled = false;
                kryptonContextMenuItem22.Enabled = false;
                kryptonContextMenuItem23.Enabled = false;
                kryptonContextMenuItem24.Enabled = false;
                kryptonContextMenuItem25.Enabled = false;
                kryptonContextMenuItem26.Enabled = false;
                kryptonContextMenuItem28.Enabled = false;
                kryptonContextMenuItem29.Enabled = false;
                if (productId == 46) // патриция 1.x
                {
                    kryptonContextMenuItem21.Enabled = true;
                }
                if (productId == 61) // куб
                {
                    kryptonContextMenuItem22.Enabled = true;
                }
                if (productId == 63) // норманн
                {
                    kryptonContextMenuItem23.Enabled = true;
                }
                if (productId == 73) // патриция 1.0
                {
                    kryptonContextMenuItem24.Enabled = true;
                }
                if (productId == 74) // патриция 2.x
                {
                    kryptonContextMenuItem25.Enabled = true;
                }
                if (productId == 75) // патриция 3.x
                {
                    kryptonContextMenuItem26.Enabled = true;
                }
                if (productId == 80) // ПАТРИЦИЯ-4.0 (С)
                {
                    kryptonContextMenuItem28.Enabled = true;
                }
                if (productId == 82) // КРИСМАР
                {
                    kryptonContextMenuItem29.Enabled = true;
                }
            }
        }

        private void kryptonContextMenuItem4_Click(object sender, EventArgs e)
        {
            int length = 0;
            int height = 0;
            int width = 0;
            int decorConfigId = 0;

            if (_decorCatalog.LengthBindingSource.Count > 0)
                length = Convert.ToInt32(((DataRowView)_decorCatalog.LengthBindingSource.Current).Row["Length"]);

            if (_decorCatalog.HeightBindingSource.Count > 0)
                height = Convert.ToInt32(((DataRowView)_decorCatalog.HeightBindingSource.Current).Row["Height"]);

            if (_decorCatalog.WidthBindingSource.Count > 0)
                width = Convert.ToInt32(((DataRowView)_decorCatalog.WidthBindingSource.Current).Row["Width"]);

            bool bNeedLength = false;
            bool bNeedHeight = false;
            bool bNeedWidth = false;
            bool pressOk = false;
            int labelsCount = 0;
            int positionsCount = 0;
            int labelsLength = 0;
            int labelsHeight = 0;
            int labelsWidth = 0;
            if (length == 0)
                bNeedLength = true;
            if (height == 0)
                bNeedHeight = true;
            if (width == 0)
                bNeedWidth = true;
            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            SampleDecorLabelInfoMenu sampleLabelInfoMenu = new SampleDecorLabelInfoMenu(this, bNeedLength, bNeedHeight, bNeedWidth);
            _topForm = sampleLabelInfoMenu;
            sampleLabelInfoMenu.ShowDialog();
            pressOk = sampleLabelInfoMenu.PressOK;
            labelsCount = sampleLabelInfoMenu.LabelsCount;
            labelsLength = sampleLabelInfoMenu.LabelsLength;
            labelsHeight = sampleLabelInfoMenu.LabelsHeight;
            labelsWidth = sampleLabelInfoMenu.LabelsWidth;
            positionsCount = sampleLabelInfoMenu.PositionsCount;
            phantomForm.Close();
            phantomForm.Dispose();
            sampleLabelInfoMenu.Dispose();
            _topForm = null;

            if (!pressOk)
                return;

            if (length > 0)
                labelsLength = length;
            if (height > 0)
                labelsHeight = height;
            if (width > 0)
                labelsWidth = width;
            string product = ProductsDataGrid.SelectedRows[0].Cells["ProductName"].Value.ToString();
            string decor = DecorDataGrid.SelectedRows[0].Cells["Name"].Value.ToString();
            string color = ColorsDataGrid.SelectedRows[0].Cells["ColorName"].Value.ToString();
            string patina = DecorPatinaDataGrid.SelectedRows[0].Cells["PatinaName"].Value.ToString();
            if (patina.Length > 0 && patina != "-")
                color += " " + patina;

            if (_decorCatalog != null)
                if (_decorCatalog.WidthBindingSource.Count > 0)
                {
                    if (((DataRowView)_decorCatalog.WidthBindingSource.Current)["Width"] != DBNull.Value)
                    {
                        decorConfigId = _decorCatalog.GetDecorConfigID(Convert.ToInt32(((DataRowView)_decorCatalog.DecorProductsBindingSource.Current)["ProductID"]),
                                                      ((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString(),
                                                      Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]),
                                                      Convert.ToInt32(((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current)["InsetTypeID"]),
                    Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current)["InsetColorID"]),
                                                      length, height, width);
                    }
                }

            DataRow newRow = _decorDt.NewRow();

            newRow["Decor"] = decor;
            newRow["Color"] = color;
            newRow["Product"] = product;
            newRow["Length"] = labelsLength.ToString();
            newRow["Height"] = labelsHeight.ToString();
            newRow["Width"] = labelsWidth.ToString();
            newRow["LabelsCount"] = labelsCount;
            newRow["PositionsCount"] = positionsCount;
            newRow["DecorConfigID"] = decorConfigId;

            if (kryptonCheckSet2.CheckedIndex == 0)
                newRow["FactoryType"] = 2;
            if (kryptonCheckSet2.CheckedIndex == 1)
                newRow["FactoryType"] = 1;
            _decorDt.Rows.Add(newRow);
            //DecorDT = DecorCatalog.GetBagetDT();
        }

        private void kryptonContextMenuItem5_Click(object sender, EventArgs e)
        {
            _sampleLabelManager.ClearOrderData();
            _decorDt.Clear();
        }

        private void kryptonContextMenuItem6_Click(object sender, EventArgs e)
        {
            _sampleLabelManager.ClearLabelInfo();

            //Проверка
            if (_decorDt.Rows.Count == 0)
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                    "Очередь для печати пуста",
                    "Ошибка печати");
                return;
            }

            for (int i = 0; i < _decorDt.Rows.Count; i++)
            {
                int labelsCount = Convert.ToInt32(_decorDt.Rows[i]["LabelsCount"]);
                int decorConfigId = Convert.ToInt32(_decorDt.Rows[i]["DecorConfigID"]);
                int sampleLabelId = 0;
                for (int j = 0; j < labelsCount; j++)
                {
                    SampleInfo labelInfo = new SampleInfo();

                    DataTable dt = _decorDt.Clone();

                    DataRow destRow = dt.NewRow();
                    foreach (DataColumn column in _decorDt.Columns)
                    {
                        if (column.ColumnName == "FactoryType")
                            continue;
                        destRow[column.ColumnName] = _decorDt.Rows[i][column.ColumnName];
                    }
                    labelInfo.ProductType = 1;
                    sampleLabelId = _sampleLabelManager.SaveSampleLabel(decorConfigId, DateTime.Now, Security.CurrentUserID, 1);
                    labelInfo.BarcodeNumber = _sampleLabelManager.GetBarcodeNumber(17, sampleLabelId);
                    labelInfo.FactoryType = Convert.ToInt32(_decorDt.Rows[i]["FactoryType"]);
                    labelInfo.DocDateTime = DateTime.Now.ToString("dd.MM.yyyy");
                    dt.Rows.Add(destRow);
                    labelInfo.OrderData = dt;

                    _sampleLabelManager.AddLabelInfo(ref labelInfo);
                }
            }
            PrintDialog.Document = _sampleLabelManager.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                _sampleLabelManager.Print();
            }
        }

        private void FrontsDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void FrameColorsDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void TechnoFrameColorsDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void PatinaDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void InsetTypesDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void InsetColorsDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void TechnoInsetTypesDataGridDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void TechnoInsetColorsDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void FrontsHeightDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void openFileDialog4_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var fileInfo = new System.IO.FileInfo(openFileDialog4.FileName);

            //if (fileInfo.Length > 1500000)
            //{
            //    MessageBox.Show("Файл больше 1,5 МБ и не может быть загружен");
            //    return;
            //}
            _attachmentsDt.Clear();

            DataRow newRow = _attachmentsDt.NewRow();
            newRow["FileName"] = Path.GetFileNameWithoutExtension(openFileDialog4.FileName);
            newRow["Extension"] = fileInfo.Extension;
            newRow["Path"] = openFileDialog4.FileName;
            _attachmentsDt.Rows.Add(newRow);

            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            string description = _frontsCatalog.GetFrontDescription(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);
            //if (Description.Length == 0)
            //{
            //    MessageBox.Show("Отсутствует Description");
            //    return;
            //}
            int configId = _frontsCatalog.SaveFrontAttachments(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);
            if (configId != -1)
                _frontsCatalog.AttachConfigImage(_attachmentsDt, configId, 0, description, frontName,
                    FrameColorsDataGrid.SelectedRows[0].Cells["ColorName"].FormattedValue.ToString(),
                    PatinaDataGrid.SelectedRows[0].Cells["PatinaName"].FormattedValue.ToString());

            pcbxFrontImage.Image = null;

            if (cbFrontImage.Checked)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                if (configId != -1)
                    pcbxFrontImage.Image = _frontsCatalog.GetFrontConfigImage(configId);

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxFrontImage.Image == null)
                    pcbxFrontImage.Cursor = Cursors.Default;
                else
                    pcbxFrontImage.Cursor = Cursors.Hand;
            }
        }

        private void cbFrontImage_CheckedChanged(object sender, EventArgs e)
        {
            if (_frontsCatalog == null)
                return;

            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            pcbxFrontImage.Image = null;

            if (cbFrontImage.Checked)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                int configId = _frontsCatalog.GetFrontAttachments(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);
                if (configId != -1)
                    pcbxFrontImage.Image = _frontsCatalog.GetFrontConfigImage(configId);

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxFrontImage.Image == null)
                    pcbxFrontImage.Cursor = Cursors.Default;
                else
                    pcbxFrontImage.Cursor = Cursors.Hand;
            }
        }

        private void pcbxFrontImage_Click(object sender, EventArgs e)
        {
            if (pcbxFrontImage.Image == null)
                return;

            PhantomForm phantomForm = new PhantomForm();
            phantomForm.Show();

            ZoomImageForm zoomImageForm = new ZoomImageForm(pcbxFrontImage.Image, ref _topForm);

            _topForm = zoomImageForm;

            zoomImageForm.ShowDialog();

            phantomForm.Close();
            phantomForm.Dispose();

            _topForm = null;

            zoomImageForm.Dispose();
        }

        private void btnAddFrontImageButton_Click(object sender, EventArgs e)
        {
            if (FrontsDataGrid.SelectedRows.Count == 0)
                return;

            openFileDialog4.ShowDialog();
        }

        private void btnRemoveFrontImageButton_Click(object sender, EventArgs e)
        {
            bool okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                   "Вы уверены, что хотите удалить?",
                   "Удаление");

            if (!okCancel)
                return;

            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            int configId = _frontsCatalog.GetFrontAttachments(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);
            if (configId != -1)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                _frontsCatalog.DetachConfigImage(configId);
                pcbxFrontImage.Image = _frontsCatalog.GetFrontConfigImage(configId);

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxFrontImage.Image == null)
                    pcbxFrontImage.Cursor = Cursors.Default;
                else
                    pcbxFrontImage.Cursor = Cursors.Hand;
            }
        }

        private void cbFrontTechStore_CheckedChanged(object sender, EventArgs e)
        {
            if (_frontsCatalog == null)
                return;

            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            pcbxFrontTechStore.Image = null;

            if (cbFrontTechStore.Checked)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                int frontId = _frontsCatalog.GetFrontID(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);
                if (frontId != -1)
                    pcbxFrontTechStore.Image = _frontsCatalog.GetTechStoreImage(frontId);

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxFrontTechStore.Image == null)
                    pcbxFrontTechStore.Cursor = Cursors.Default;
                else
                    pcbxFrontTechStore.Cursor = Cursors.Hand;
            }
        }

        private void pcbxFrontTechStore_Click(object sender, EventArgs e)
        {
            if (pcbxFrontTechStore.Image == null)
                return;

            PhantomForm phantomForm = new PhantomForm();
            phantomForm.Show();

            ZoomImageForm zoomImageForm = new ZoomImageForm(pcbxFrontTechStore.Image, ref _topForm);

            _topForm = zoomImageForm;

            zoomImageForm.ShowDialog();

            phantomForm.Close();
            phantomForm.Dispose();

            _topForm = null;

            zoomImageForm.Dispose();
        }

        private void btnAddDecorImageButton_Click(object sender, EventArgs e)
        {
            if (ProductsDataGrid.SelectedRows.Count == 0)
                return;

            openFileDialog1.ShowDialog();
        }

        private void btnRemoveDecorImageButton_Click(object sender, EventArgs e)
        {
            bool okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                   "Вы уверены, что хотите удалить?",
                   "Удаление");

            if (!okCancel)
                return;

            int productId = Convert.ToInt32(((DataRowView)_decorCatalog.DecorProductsBindingSource.Current)["ProductID"]);
            string decorName = ((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString();
            int colorId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]);
            int patinaId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaID"]);
            int insetTypeId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current)["InsetTypeID"]);
            int insetColorId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current)["InsetColorID"]);

            int configId = _decorCatalog.GetDecorAttachments(productId, decorName, colorId, patinaId, insetTypeId, insetColorId);
            if (configId != -1)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                _decorCatalog.DetachConfigImage(configId);
                pcbxDecorImage.Image = _decorCatalog.GetDecorConfigImage(configId, productId);

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxDecorImage.Image == null)
                    pcbxDecorImage.Cursor = Cursors.Default;
                else
                    pcbxDecorImage.Cursor = Cursors.Hand;
            }
        }

        private void pcbxDecorImage_Click(object sender, EventArgs e)
        {
            if (pcbxDecorImage.Image == null)
                return;

            PhantomForm phantomForm = new PhantomForm();
            phantomForm.Show();

            ZoomImageForm zoomImageForm = new ZoomImageForm(pcbxDecorImage.Image, ref _topForm);

            _topForm = zoomImageForm;

            zoomImageForm.ShowDialog();

            phantomForm.Close();
            phantomForm.Dispose();

            _topForm = null;

            zoomImageForm.Dispose();
        }

        private void pcbxDecorTechStore_Click(object sender, EventArgs e)
        {
            if (pcbxDecorTechStore.Image == null)
                return;

            PhantomForm phantomForm = new PhantomForm();
            phantomForm.Show();

            ZoomImageForm zoomImageForm = new ZoomImageForm(pcbxDecorTechStore.Image, ref _topForm);

            _topForm = zoomImageForm;

            zoomImageForm.ShowDialog();

            phantomForm.Close();
            phantomForm.Dispose();

            _topForm = null;

            zoomImageForm.Dispose();
        }

        private void cbDecorImage_CheckedChanged(object sender, EventArgs e)
        {
            if (_decorCatalog == null)
                return;

            int productId = Convert.ToInt32(((DataRowView)_decorCatalog.DecorProductsBindingSource.Current)["ProductID"]);
            string decorName = ((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString();
            int colorId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]);
            int patinaId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaID"]);
            int insetTypeId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current)["InsetTypeID"]);
            int insetColorId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current)["InsetColorID"]);

            pcbxDecorImage.Image = null;

            if (cbDecorImage.Checked)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                int configId = _decorCatalog.GetDecorAttachments(productId, decorName, colorId, patinaId, insetTypeId, insetColorId);
                if (configId != -1)
                    pcbxDecorImage.Image = _decorCatalog.GetDecorConfigImage(configId, productId);

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxDecorImage.Image == null)
                    pcbxDecorImage.Cursor = Cursors.Default;
                else
                    pcbxDecorImage.Cursor = Cursors.Hand;
            }
        }

        private void cbDecorTechStore_CheckedChanged(object sender, EventArgs e)
        {
            if (_decorCatalog == null)
                return;

            int productId = Convert.ToInt32(((DataRowView)_decorCatalog.DecorProductsBindingSource.Current)["ProductID"]);
            string decorName = ((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString();
            int colorId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]);
            int patinaId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaID"]);
            int insetTypeId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current)["InsetTypeID"]);
            int insetColorId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current)["InsetColorID"]);

            pcbxDecorTechStore.Image = null;

            if (cbDecorTechStore.Checked)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                int decorId = _decorCatalog.GetDecorID(productId, decorName, colorId, patinaId, insetTypeId, insetColorId);
                if (decorId != -1)
                    pcbxDecorTechStore.Image = _decorCatalog.GetTechStoreImage(decorId);

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxDecorTechStore.Image == null)
                    pcbxDecorTechStore.Cursor = Cursors.Default;
                else
                    pcbxDecorTechStore.Cursor = Cursors.Hand;
            }
        }

        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var fileInfo = new System.IO.FileInfo(openFileDialog1.FileName);

            //if (fileInfo.Length > 1500000)
            //{
            //    MessageBox.Show("Файл больше 1,5 МБ и не может быть загружен");
            //    return;
            //}
            _attachmentsDt.Clear();

            DataRow newRow = _attachmentsDt.NewRow();
            newRow["FileName"] = Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
            newRow["Extension"] = fileInfo.Extension;
            newRow["Path"] = openFileDialog1.FileName;
            _attachmentsDt.Rows.Add(newRow);

            int productId = Convert.ToInt32(((DataRowView)_decorCatalog.DecorProductsBindingSource.Current)["ProductID"]);
            string decorName = ((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString();
            int colorId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]);
            int patinaId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaID"]);
            int insetTypeId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current)["InsetTypeID"]);
            int insetColorId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current)["InsetColorID"]);
            int productType = 1;

            //if (ProductID == 46 || ProductID == 61 || ProductID == 62 || ProductID == 63)
            if (CheckOrdersStatus.IsCabFurniture(productId))
                productType = 2;
            int configId = _decorCatalog.SaveDecorAttachments(productId, decorName, colorId, patinaId, insetTypeId, insetColorId);
            if (configId != -1)
                _decorCatalog.AttachConfigImage(_attachmentsDt, configId, productType, ((DataRowView)_decorCatalog.DecorProductsBindingSource.Current)["ProductName"].ToString(),
                    decorName, ((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorName"].ToString(),
                    ((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaName"].ToString(),
                    ((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current)["InsetColorName"].ToString());

            pcbxDecorImage.Image = null;
            if (cbDecorImage.Checked)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();
                while (!SplashWindow.bSmallCreated) ;

                if (configId != -1)
                    pcbxDecorImage.Image = _decorCatalog.GetDecorConfigImage(configId, productId);

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxDecorImage.Image == null)
                    pcbxDecorImage.Cursor = Cursors.Default;
                else
                    pcbxDecorImage.Cursor = Cursors.Hand;
            }
        }

        private void FrontsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            FrameColorsDataGrid.SelectionChanged -= FrameColorsDataGrid_SelectionChanged;
            TechnoFrameColorsDataGrid.SelectionChanged -= TechnoFrameColorsDataGrid_SelectionChanged;
            InsetTypesDataGrid.SelectionChanged -= InsetTypesDataGrid_SelectionChanged;
            InsetColorsDataGrid.SelectionChanged -= InsetColorsDataGrid_SelectionChanged;
            TechnoInsetTypesDataGrid.SelectionChanged -= TechnoInsetTypesDataGrid_SelectionChanged;
            TechnoInsetColorsDataGrid.SelectionChanged -= TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged -= PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged -= FrontsHeightDataGrid_SelectionChanged;
            GetFrameColors();
            GetTechnoFrameColors();
            GetInsetTypes();
            GetInsetColors();
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetPatina();
            GetFrontHeight();
            GetFrontWidth();
            FrameColorsDataGrid.SelectionChanged += FrameColorsDataGrid_SelectionChanged;
            TechnoFrameColorsDataGrid.SelectionChanged += TechnoFrameColorsDataGrid_SelectionChanged;
            InsetTypesDataGrid.SelectionChanged += InsetTypesDataGrid_SelectionChanged;
            InsetColorsDataGrid.SelectionChanged += InsetColorsDataGrid_SelectionChanged;
            TechnoInsetTypesDataGrid.SelectionChanged += TechnoInsetTypesDataGrid_SelectionChanged;
            TechnoInsetColorsDataGrid.SelectionChanged += TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged += PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged += FrontsHeightDataGrid_SelectionChanged;
        }

        private void FrameColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            TechnoFrameColorsDataGrid.SelectionChanged -= TechnoFrameColorsDataGrid_SelectionChanged;
            InsetTypesDataGrid.SelectionChanged -= InsetTypesDataGrid_SelectionChanged;
            InsetColorsDataGrid.SelectionChanged -= InsetColorsDataGrid_SelectionChanged;
            TechnoInsetTypesDataGrid.SelectionChanged -= TechnoInsetTypesDataGrid_SelectionChanged;
            TechnoInsetColorsDataGrid.SelectionChanged -= TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged -= PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged -= FrontsHeightDataGrid_SelectionChanged;
            GetTechnoFrameColors();
            GetInsetTypes();
            GetInsetColors();
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetPatina();
            GetFrontHeight();
            GetFrontWidth();
            TechnoFrameColorsDataGrid.SelectionChanged += TechnoFrameColorsDataGrid_SelectionChanged;
            InsetTypesDataGrid.SelectionChanged += InsetTypesDataGrid_SelectionChanged;
            InsetColorsDataGrid.SelectionChanged += InsetColorsDataGrid_SelectionChanged;
            TechnoInsetTypesDataGrid.SelectionChanged += TechnoInsetTypesDataGrid_SelectionChanged;
            TechnoInsetColorsDataGrid.SelectionChanged += TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged += PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged += FrontsHeightDataGrid_SelectionChanged;
        }

        private void TechnoFrameColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            InsetTypesDataGrid.SelectionChanged -= InsetTypesDataGrid_SelectionChanged;
            InsetColorsDataGrid.SelectionChanged -= InsetColorsDataGrid_SelectionChanged;
            TechnoInsetTypesDataGrid.SelectionChanged -= TechnoInsetTypesDataGrid_SelectionChanged;
            TechnoInsetColorsDataGrid.SelectionChanged -= TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged -= PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged -= FrontsHeightDataGrid_SelectionChanged;
            GetInsetTypes();
            GetInsetColors();
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetPatina();
            GetFrontHeight();
            GetFrontWidth();
            InsetTypesDataGrid.SelectionChanged += InsetTypesDataGrid_SelectionChanged;
            InsetColorsDataGrid.SelectionChanged += InsetColorsDataGrid_SelectionChanged;
            TechnoInsetTypesDataGrid.SelectionChanged += TechnoInsetTypesDataGrid_SelectionChanged;
            TechnoInsetColorsDataGrid.SelectionChanged += TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged += PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged += FrontsHeightDataGrid_SelectionChanged;
        }

        private void InsetTypesDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            InsetColorsDataGrid.SelectionChanged -= InsetColorsDataGrid_SelectionChanged;
            TechnoInsetTypesDataGrid.SelectionChanged -= TechnoInsetTypesDataGrid_SelectionChanged;
            TechnoInsetColorsDataGrid.SelectionChanged -= TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged -= PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged -= FrontsHeightDataGrid_SelectionChanged;
            GetInsetColors();
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetPatina();
            GetFrontHeight();
            GetFrontWidth();
            InsetColorsDataGrid.SelectionChanged += InsetColorsDataGrid_SelectionChanged;
            TechnoInsetTypesDataGrid.SelectionChanged += TechnoInsetTypesDataGrid_SelectionChanged;
            TechnoInsetColorsDataGrid.SelectionChanged += TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged += PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged += FrontsHeightDataGrid_SelectionChanged;
        }

        private void InsetColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            TechnoInsetTypesDataGrid.SelectionChanged -= TechnoInsetTypesDataGrid_SelectionChanged;
            TechnoInsetColorsDataGrid.SelectionChanged -= TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged -= PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged -= FrontsHeightDataGrid_SelectionChanged;
            GetTechnoInsetTypes();
            GetTechnoInsetColors();
            GetPatina();
            GetFrontHeight();
            GetFrontWidth();
            TechnoInsetTypesDataGrid.SelectionChanged += TechnoInsetTypesDataGrid_SelectionChanged;
            TechnoInsetColorsDataGrid.SelectionChanged += TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged += PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged += FrontsHeightDataGrid_SelectionChanged;
        }

        private void TechnoInsetTypesDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            TechnoInsetColorsDataGrid.SelectionChanged -= TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged -= PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged -= FrontsHeightDataGrid_SelectionChanged;
            GetTechnoInsetColors();
            GetPatina();
            GetFrontHeight();
            GetFrontWidth();
            TechnoInsetColorsDataGrid.SelectionChanged += TechnoInsetColorsDataGrid_SelectionChanged;
            PatinaDataGrid.SelectionChanged += PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged += FrontsHeightDataGrid_SelectionChanged;
        }

        private void TechnoInsetColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            PatinaDataGrid.SelectionChanged -= PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged -= FrontsHeightDataGrid_SelectionChanged;
            GetPatina();
            GetFrontHeight();
            GetFrontWidth();
            PatinaDataGrid.SelectionChanged += PatinaDataGrid_SelectionChanged;
            FrontsHeightDataGrid.SelectionChanged += FrontsHeightDataGrid_SelectionChanged;
        }

        private void PatinaDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            FrontsHeightDataGrid.SelectionChanged -= FrontsHeightDataGrid_SelectionChanged;
            GetFrontHeight();
            GetFrontWidth();
            FrontsHeightDataGrid.SelectionChanged += FrontsHeightDataGrid_SelectionChanged;
        }

        private void FrontsHeightDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            GetFrontWidth();
        }

        private void ProductsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            DecorDataGrid.SelectionChanged -= DecorDataGrid_SelectionChanged;
            ColorsDataGrid.SelectionChanged -= ColorsDataGrid_SelectionChanged;
            DecorPatinaDataGrid.SelectionChanged -= DecorPatinaDataGrid_SelectionChanged;
            DecorInsetTypesDataGrid.SelectionChanged -= DecorInsetTypesDataGrid_SelectionChanged;
            DecorInsetColorsDataGrid.SelectionChanged -= DecorInsetColorsDataGrid_SelectionChanged;
            LengthDataGrid.SelectionChanged -= LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged -= HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged -= WidthDataGrid_SelectionChanged;
            GetDecorItems();
            GetDecorColors();
            GetDecorPatina();
            GetDecorInsetTypes();
            GetDecorInsetColors();
            GetDecorLength();
            GetDecorHeight();
            GetDecorWidth();
            DecorDataGrid.SelectionChanged += DecorDataGrid_SelectionChanged;
            ColorsDataGrid.SelectionChanged += ColorsDataGrid_SelectionChanged;
            DecorPatinaDataGrid.SelectionChanged += DecorPatinaDataGrid_SelectionChanged;
            DecorInsetTypesDataGrid.SelectionChanged += DecorInsetTypesDataGrid_SelectionChanged;
            DecorInsetColorsDataGrid.SelectionChanged += DecorInsetColorsDataGrid_SelectionChanged;
            LengthDataGrid.SelectionChanged += LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged += HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged += WidthDataGrid_SelectionChanged;
        }

        private void DecorDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            ColorsDataGrid.SelectionChanged -= ColorsDataGrid_SelectionChanged;
            DecorPatinaDataGrid.SelectionChanged -= DecorPatinaDataGrid_SelectionChanged;
            DecorInsetTypesDataGrid.SelectionChanged -= DecorInsetTypesDataGrid_SelectionChanged;
            DecorInsetColorsDataGrid.SelectionChanged -= DecorInsetColorsDataGrid_SelectionChanged;
            LengthDataGrid.SelectionChanged -= LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged -= HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged -= WidthDataGrid_SelectionChanged;
            GetDecorColors();
            GetDecorPatina();
            GetDecorInsetTypes();
            GetDecorInsetColors();
            GetDecorLength();
            GetDecorHeight();
            GetDecorWidth();
            ColorsDataGrid.SelectionChanged += ColorsDataGrid_SelectionChanged;
            DecorPatinaDataGrid.SelectionChanged += DecorPatinaDataGrid_SelectionChanged;
            DecorInsetTypesDataGrid.SelectionChanged += DecorInsetTypesDataGrid_SelectionChanged;
            DecorInsetColorsDataGrid.SelectionChanged += DecorInsetColorsDataGrid_SelectionChanged;
            LengthDataGrid.SelectionChanged += LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged += HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged += WidthDataGrid_SelectionChanged;
        }

        private void ColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            DecorPatinaDataGrid.SelectionChanged -= DecorPatinaDataGrid_SelectionChanged;
            DecorInsetTypesDataGrid.SelectionChanged -= DecorInsetTypesDataGrid_SelectionChanged;
            DecorInsetColorsDataGrid.SelectionChanged -= DecorInsetColorsDataGrid_SelectionChanged;
            LengthDataGrid.SelectionChanged -= LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged -= HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged -= WidthDataGrid_SelectionChanged;
            GetDecorPatina();
            GetDecorInsetTypes();
            GetDecorInsetColors();
            GetDecorLength();
            GetDecorHeight();
            GetDecorWidth();
            DecorPatinaDataGrid.SelectionChanged += DecorPatinaDataGrid_SelectionChanged;
            DecorInsetTypesDataGrid.SelectionChanged += DecorInsetTypesDataGrid_SelectionChanged;
            DecorInsetColorsDataGrid.SelectionChanged += DecorInsetColorsDataGrid_SelectionChanged;
            LengthDataGrid.SelectionChanged += LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged += HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged += WidthDataGrid_SelectionChanged;
        }

        private void DecorPatinaDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            DecorInsetTypesDataGrid.SelectionChanged -= DecorInsetTypesDataGrid_SelectionChanged;
            DecorInsetColorsDataGrid.SelectionChanged -= DecorInsetColorsDataGrid_SelectionChanged;
            LengthDataGrid.SelectionChanged -= LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged -= HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged -= WidthDataGrid_SelectionChanged;
            GetDecorInsetTypes();
            GetDecorInsetColors();
            GetDecorLength();
            GetDecorHeight();
            GetDecorWidth();
            DecorInsetTypesDataGrid.SelectionChanged += DecorInsetTypesDataGrid_SelectionChanged;
            DecorInsetColorsDataGrid.SelectionChanged += DecorInsetColorsDataGrid_SelectionChanged;
            LengthDataGrid.SelectionChanged += LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged += HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged += WidthDataGrid_SelectionChanged;
        }

        private void DecorInsetTypesDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            DecorInsetColorsDataGrid.SelectionChanged -= DecorInsetColorsDataGrid_SelectionChanged;
            LengthDataGrid.SelectionChanged -= LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged -= HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged -= WidthDataGrid_SelectionChanged;
            GetDecorInsetColors();
            GetDecorLength();
            GetDecorHeight();
            GetDecorWidth();
            DecorInsetColorsDataGrid.SelectionChanged += DecorInsetColorsDataGrid_SelectionChanged;
            LengthDataGrid.SelectionChanged += LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged += HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged += WidthDataGrid_SelectionChanged;
        }


        private void DecorInsetColorsDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            LengthDataGrid.SelectionChanged -= LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged -= HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged -= WidthDataGrid_SelectionChanged;
            GetDecorLength();
            GetDecorHeight();
            GetDecorWidth();
            LengthDataGrid.SelectionChanged += LengthDataGrid_SelectionChanged;
            HeightDataGrid.SelectionChanged += HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged += WidthDataGrid_SelectionChanged;
        }

        private void LengthDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            HeightDataGrid.SelectionChanged -= HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged -= WidthDataGrid_SelectionChanged;
            GetDecorHeight();
            GetDecorWidth();
            HeightDataGrid.SelectionChanged += HeightDataGrid_SelectionChanged;
            WidthDataGrid.SelectionChanged += WidthDataGrid_SelectionChanged;
        }

        private void HeightDataGrid_SelectionChanged(object sender, EventArgs e)
        {
            WidthDataGrid.SelectionChanged -= WidthDataGrid_SelectionChanged;
            GetDecorWidth();
            WidthDataGrid.SelectionChanged += WidthDataGrid_SelectionChanged;
        }

        private void WidthDataGrid_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void kryptonContextMenuItem13_Click(object sender, EventArgs e)
        {
            if (InsetColorsDataGrid.SelectedRows.Count == 0)
                return;

            openFileDialog3.ShowDialog();
        }

        private void kryptonContextMenuItem15_Click(object sender, EventArgs e)
        {
            bool okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                      "Вы уверены, что хотите удалить?",
                      "Удаление");

            if (!okCancel)
                return;

            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            int configId = _frontsCatalog.GetInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
            _frontsCatalog.DetachInsetColorImage(configId);
            pcbxVisualConfig.Image = null;
            if (cbVisualConfig.Checked)
            {
                _layer = null;
                _backgr = null;
                _backgr2 = null;
                _newBitmap = null;
                _newBitmap2 = null;
                configId = _frontsCatalog.GetInsetTypesConfig(frontName, colorId, technoColorId);
                Image img = _frontsCatalog.GetInsetTypeImage(configId);
                if (configId != -1 && img != null)
                    _layer = new Bitmap(img);
                configId = _frontsCatalog.GetInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr = new Bitmap(img);

                configId = _frontsCatalog.GetTechnoInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetTechnoInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr2 = new Bitmap(img);

                if (_layer != null)
                {
                    if (_backgr != null)
                    {
                        if (_backgr2 != null)
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            int width2 = _backgr2.Width;
                            int height2 = _backgr2.Height;
                            _newBitmap = new Bitmap(width, height);
                            _newBitmap2 = new Bitmap(width1, height1);
                            using (var canvas = Graphics.FromImage(_newBitmap2))
                            {
                                if (height2 >= height1)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr2.Width, _backgr2.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }

                                int width3 = _newBitmap2.Width;
                                int height3 = _newBitmap2.Height;
                                using (var canvas1 = Graphics.FromImage(_newBitmap))
                                {
                                    if (height3 >= height)
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                    else
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _newBitmap2.Width, _newBitmap2.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                        else
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            _newBitmap = new Bitmap(width, height);
                            using (var canvas = Graphics.FromImage(_newBitmap))
                            {
                                if (height1 >= height)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                    }
                    else
                        pcbxVisualConfig.Image = _layer;
                }
                else
                {
                    if (_backgr != null)
                        pcbxVisualConfig.Image = _backgr;
                }

                if (pcbxVisualConfig.Image == null)
                    pcbxVisualConfig.Cursor = Cursors.Default;
                else
                    pcbxVisualConfig.Cursor = Cursors.Hand;
            }

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void kryptonContextMenuItem14_Click(object sender, EventArgs e)
        {
            bool okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                      "Вы уверены, что хотите удалить?",
                      "Удаление");

            if (!okCancel)
                return;

            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            int configId = _frontsCatalog.GetInsetTypesConfig(frontName, colorId, technoColorId);
            _frontsCatalog.DetachInsetTypeImage(configId);
            pcbxVisualConfig.Image = null;
            if (cbVisualConfig.Checked)
            {
                _layer = null;
                _backgr = null;
                _backgr2 = null;
                _newBitmap = null;
                _newBitmap2 = null;
                Image img = _frontsCatalog.GetInsetTypeImage(configId);
                if (configId != -1 && img != null)
                    _layer = new Bitmap(img);
                configId = _frontsCatalog.GetInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr = new Bitmap(img);

                configId = _frontsCatalog.GetTechnoInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetTechnoInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr2 = new Bitmap(img);

                if (_layer != null)
                {
                    if (_backgr != null)
                    {
                        if (_backgr2 != null)
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            int width2 = _backgr2.Width;
                            int height2 = _backgr2.Height;
                            _newBitmap = new Bitmap(width, height);
                            _newBitmap2 = new Bitmap(width1, height1);
                            using (var canvas = Graphics.FromImage(_newBitmap2))
                            {
                                if (height2 >= height1)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr2.Width, _backgr2.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }

                                int width3 = _newBitmap2.Width;
                                int height3 = _newBitmap2.Height;
                                using (var canvas1 = Graphics.FromImage(_newBitmap))
                                {
                                    if (height3 >= height)
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                    else
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _newBitmap2.Width, _newBitmap2.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                        else
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            _newBitmap = new Bitmap(width, height);
                            using (var canvas = Graphics.FromImage(_newBitmap))
                            {
                                if (height1 >= height)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                    }
                    else
                        pcbxVisualConfig.Image = _layer;
                }
                else
                {
                    if (_backgr != null)
                        pcbxVisualConfig.Image = _backgr;
                }

                if (pcbxVisualConfig.Image == null)
                    pcbxVisualConfig.Cursor = Cursors.Default;
                else
                    pcbxVisualConfig.Cursor = Cursors.Hand;


                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void kryptonContextMenuItem12_Click(object sender, EventArgs e)
        {
            if (InsetTypesDataGrid.SelectedRows.Count == 0)
                return;

            openFileDialog2.ShowDialog();
        }

        private void openFileDialog2_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var fileInfo = new System.IO.FileInfo(openFileDialog2.FileName);

            //if (fileInfo.Length > 1500000)
            //{
            //    MessageBox.Show("Файл больше 1,5 МБ и не может быть загружен");
            //    return;
            //}
            _attachmentsDt.Clear();

            DataRow newRow = _attachmentsDt.NewRow();
            newRow["FileName"] = Path.GetFileNameWithoutExtension(openFileDialog2.FileName);
            newRow["Extension"] = fileInfo.Extension;
            newRow["Path"] = openFileDialog2.FileName;
            _attachmentsDt.Rows.Add(newRow);

            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            int configId = _frontsCatalog.SaveInsetTypesConfig(frontName, colorId, technoColorId);
            if (configId != -1)
                _frontsCatalog.AttachInsetTypeImage(_attachmentsDt, configId);

            pcbxVisualConfig.Image = null;
            if (cbVisualConfig.Checked)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                _layer = null;
                _backgr = null;
                _newBitmap = null;
                Image img = _frontsCatalog.GetInsetTypeImage(configId);
                if (configId != -1 && img != null)
                    _layer = new Bitmap(img);
                configId = _frontsCatalog.GetInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr = new Bitmap(img);
                if (_layer != null)
                {
                    if (_backgr != null)
                    {
                        int width = _layer.Width;
                        int height = _layer.Height;
                        int width1 = _backgr.Width;
                        int height1 = _backgr.Height;
                        _newBitmap = new Bitmap(width, height);
                        using (var canvas = Graphics.FromImage(_newBitmap))
                        {
                            if (height1 >= height)
                            {
                                canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                canvas.Save();
                            }
                            else
                            {
                                canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                canvas.Save();
                            }
                        }
                        try
                        {
                            pcbxVisualConfig.Image = _newBitmap;
                        }
                        catch (Exception ex) { MessageBox.Show(ex.Message); }
                    }
                    else
                        pcbxVisualConfig.Image = _layer;
                }
                else
                {
                    if (_backgr != null)
                        pcbxVisualConfig.Image = _backgr;
                }

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxVisualConfig.Image == null)
                    pcbxVisualConfig.Cursor = Cursors.Default;
                else
                    pcbxVisualConfig.Cursor = Cursors.Hand;
            }
        }

        private void openFileDialog3_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var fileInfo = new System.IO.FileInfo(openFileDialog3.FileName);

            _attachmentsDt.Clear();

            DataRow newRow = _attachmentsDt.NewRow();
            newRow["FileName"] = Path.GetFileNameWithoutExtension(openFileDialog3.FileName);
            newRow["Extension"] = fileInfo.Extension;
            newRow["Path"] = openFileDialog3.FileName;
            _attachmentsDt.Rows.Add(newRow);
            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            int configId = _frontsCatalog.SaveInsetColorsConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
            if (configId != -1)
                _frontsCatalog.AttachInsetColorImage(_attachmentsDt, configId);

            pcbxVisualConfig.Image = null;
            if (cbVisualConfig.Checked)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                _layer = null;
                _backgr = null;
                _backgr2 = null;
                _newBitmap = null;
                _newBitmap2 = null;
                configId = _frontsCatalog.GetInsetTypesConfig(frontName, colorId, technoColorId);
                Image img = _frontsCatalog.GetInsetTypeImage(configId);
                if (configId != -1 && img != null)
                    _layer = new Bitmap(img);
                configId = _frontsCatalog.GetInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr = new Bitmap(img);

                configId = _frontsCatalog.GetTechnoInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetTechnoInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr2 = new Bitmap(img);

                if (_layer != null)
                {
                    if (_backgr != null)
                    {
                        if (_backgr2 != null)
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            int width2 = _backgr2.Width;
                            int height2 = _backgr2.Height;
                            _newBitmap = new Bitmap(width, height);
                            _newBitmap2 = new Bitmap(width1, height1);
                            using (var canvas = Graphics.FromImage(_newBitmap2))
                            {
                                if (height2 >= height1)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr2.Width, _backgr2.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }

                                int width3 = _newBitmap2.Width;
                                int height3 = _newBitmap2.Height;
                                using (var canvas1 = Graphics.FromImage(_newBitmap))
                                {
                                    if (height3 >= height)
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                    else
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _newBitmap2.Width, _newBitmap2.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                        else
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            _newBitmap = new Bitmap(width, height);
                            using (var canvas = Graphics.FromImage(_newBitmap))
                            {
                                if (height1 >= height)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                    }
                    else
                        pcbxVisualConfig.Image = _layer;
                }
                else
                {
                    if (_backgr != null)
                        pcbxVisualConfig.Image = _backgr;
                }

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxVisualConfig.Image == null)
                    pcbxVisualConfig.Cursor = Cursors.Default;
                else
                    pcbxVisualConfig.Cursor = Cursors.Hand;
            }
        }

        private void pcbxVisualConfig_Click(object sender, EventArgs e)
        {
            if (pcbxVisualConfig.Image == null)
                return;

            PhantomForm phantomForm = new PhantomForm();
            phantomForm.Show();

            ZoomImageForm zoomImageForm = new ZoomImageForm(pcbxVisualConfig.Image, ref _topForm);

            _topForm = zoomImageForm;

            zoomImageForm.ShowDialog();

            phantomForm.Close();
            phantomForm.Dispose();

            _topForm = null;

            zoomImageForm.Dispose();
        }

        private void cbVisualConfig_CheckedChanged(object sender, EventArgs e)
        {
            if (_frontsCatalog == null)
                return;
            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            pcbxVisualConfig.Image = null;

            if (cbVisualConfig.Checked)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                _layer = null;
                _backgr = null;
                _backgr2 = null;
                _newBitmap = null;
                _newBitmap2 = null;
                int configId = _frontsCatalog.GetInsetTypesConfig(frontName, colorId, technoColorId);
                Image img = _frontsCatalog.GetInsetTypeImage(configId);
                if (configId != -1 && img != null)
                    _layer = new Bitmap(img);
                configId = _frontsCatalog.GetInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr = new Bitmap(img);

                configId = _frontsCatalog.GetTechnoInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetTechnoInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr2 = new Bitmap(img);

                if (_layer != null)
                {
                    if (_backgr != null)
                    {
                        if (_backgr2 != null)
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            int width2 = _backgr2.Width;
                            int height2 = _backgr2.Height;
                            _newBitmap = new Bitmap(width, height);
                            _newBitmap2 = new Bitmap(width1, height1);
                            using (var canvas = Graphics.FromImage(_newBitmap2))
                            {
                                if (height2 >= height1)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr2.Width, _backgr2.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }

                                int width3 = _newBitmap2.Width;
                                int height3 = _newBitmap2.Height;
                                using (var canvas1 = Graphics.FromImage(_newBitmap))
                                {
                                    if (height3 >= height)
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                    else
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _newBitmap2.Width, _newBitmap2.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                        else
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            _newBitmap = new Bitmap(width, height);
                            using (var canvas = Graphics.FromImage(_newBitmap))
                            {
                                if (height1 >= height)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                    }
                    else
                        pcbxVisualConfig.Image = _layer;
                }
                else
                {
                    if (_backgr != null)
                        pcbxVisualConfig.Image = _backgr;
                }

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxVisualConfig.Image == null)
                    pcbxVisualConfig.Cursor = Cursors.Default;
                else
                    pcbxVisualConfig.Cursor = Cursors.Hand;
            }
        }

        private void kryptonContextMenuItem16_Click(object sender, EventArgs e)
        {
            if (TechnoFrameColorsDataGrid.SelectedRows.Count == 0)
                return;

            openFileDialog5.ShowDialog();
        }

        private void TechnoFrameColorsDataGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (_bCanEdit && e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                kryptonContextMenu3.Show(new Point(Cursor.Position.X - 212, Cursor.Position.Y - 10));
            }
        }

        private void openFileDialog5_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var fileInfo = new System.IO.FileInfo(openFileDialog5.FileName);

            //if (fileInfo.Length > 1500000)
            //{
            //    MessageBox.Show("Файл больше 1,5 МБ и не может быть загружен");
            //    return;
            //}
            _attachmentsDt.Clear();

            DataRow newRow = _attachmentsDt.NewRow();
            newRow["FileName"] = Path.GetFileNameWithoutExtension(openFileDialog5.FileName);
            newRow["Extension"] = fileInfo.Extension;
            newRow["Path"] = openFileDialog5.FileName;
            _attachmentsDt.Rows.Add(newRow);

            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            int configId = _frontsCatalog.SaveInsetTypesConfig(frontName, colorId, technoColorId);
            if (configId != -1)
                _frontsCatalog.AttachInsetTypeImage(_attachmentsDt, configId);

            pcbxVisualConfig.Image = null;
            if (cbVisualConfig.Checked)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                _layer = null;
                _backgr = null;
                _newBitmap = null;
                Image img = _frontsCatalog.GetInsetTypeImage(configId);
                if (configId != -1 && img != null)
                    _layer = new Bitmap(img);
                configId = _frontsCatalog.GetInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr = new Bitmap(img);
                if (_layer != null)
                {
                    if (_backgr != null)
                    {
                        int width = _layer.Width;
                        int height = _layer.Height;
                        int width1 = _backgr.Width;
                        int height1 = _backgr.Height;
                        _newBitmap = new Bitmap(width, height);
                        using (var canvas = Graphics.FromImage(_newBitmap))
                        {
                            if (height1 >= height)
                            {
                                canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                canvas.Save();
                            }
                            else
                            {
                                canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                canvas.Save();
                            }
                        }
                        try
                        {
                            pcbxVisualConfig.Image = _newBitmap;
                        }
                        catch (Exception ex) { MessageBox.Show(ex.Message); }
                    }
                    else
                        pcbxVisualConfig.Image = _layer;
                }
                else
                {
                    if (_backgr != null)
                        pcbxVisualConfig.Image = _backgr;
                }

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxVisualConfig.Image == null)
                    pcbxVisualConfig.Cursor = Cursors.Default;
                else
                    pcbxVisualConfig.Cursor = Cursors.Hand;
            }
        }

        private void kryptonContextMenuItem17_Click(object sender, EventArgs e)
        {
            bool okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                         "Вы уверены, что хотите удалить?",
                         "Удаление");

            if (!okCancel)
                return;

            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            int configId = _frontsCatalog.GetInsetTypesConfig(frontName, colorId, technoColorId);
            _frontsCatalog.DetachInsetTypeImage(configId);
            pcbxVisualConfig.Image = null;
            if (cbVisualConfig.Checked)
            {
                _layer = null;
                _backgr = null;
                _backgr2 = null;
                _newBitmap = null;
                _newBitmap2 = null;
                Image img = _frontsCatalog.GetInsetTypeImage(configId);
                if (configId != -1 && img != null)
                    _layer = new Bitmap(img);
                configId = _frontsCatalog.GetInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr = new Bitmap(img);

                configId = _frontsCatalog.GetTechnoInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetTechnoInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr2 = new Bitmap(img);

                if (_layer != null)
                {
                    if (_backgr != null)
                    {
                        if (_backgr2 != null)
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            int width2 = _backgr2.Width;
                            int height2 = _backgr2.Height;
                            _newBitmap = new Bitmap(width, height);
                            _newBitmap2 = new Bitmap(width1, height1);
                            using (var canvas = Graphics.FromImage(_newBitmap2))
                            {
                                if (height2 >= height1)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr2.Width, _backgr2.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }

                                int width3 = _newBitmap2.Width;
                                int height3 = _newBitmap2.Height;
                                using (var canvas1 = Graphics.FromImage(_newBitmap))
                                {
                                    if (height3 >= height)
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                    else
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _newBitmap2.Width, _newBitmap2.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                        else
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            _newBitmap = new Bitmap(width, height);
                            using (var canvas = Graphics.FromImage(_newBitmap))
                            {
                                if (height1 >= height)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                    }
                    else
                        pcbxVisualConfig.Image = _layer;
                }
                else
                {
                    if (_backgr != null)
                        pcbxVisualConfig.Image = _backgr;
                }

                if (pcbxVisualConfig.Image == null)
                    pcbxVisualConfig.Cursor = Cursors.Default;
                else
                    pcbxVisualConfig.Cursor = Cursors.Hand;

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
        }

        private void kryptonContextMenuItem18_Click(object sender, EventArgs e)
        {
            if (TechnoInsetColorsDataGrid.SelectedRows.Count == 0)
                return;

            openFileDialog6.ShowDialog();
        }

        private void kryptonContextMenuItem19_Click(object sender, EventArgs e)
        {
            bool okCancel = Infinium.LightMessageBox.Show(ref _topForm, true,
                         "Вы уверены, что хотите удалить?",
                         "Удаление");

            if (!okCancel)
                return;

            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            int configId = _frontsCatalog.GetTechnoInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
            _frontsCatalog.DetachTechnoInsetColorImage(configId);
            pcbxVisualConfig.Image = null;
            if (cbVisualConfig.Checked)
            {
                _layer = null;
                _backgr = null;
                _backgr2 = null;
                _newBitmap = null;
                _newBitmap2 = null;
                configId = _frontsCatalog.GetInsetTypesConfig(frontName, colorId, technoColorId);
                Image img = _frontsCatalog.GetInsetTypeImage(configId);
                if (configId != -1 && img != null)
                    _layer = new Bitmap(img);
                configId = _frontsCatalog.GetInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr = new Bitmap(img);

                configId = _frontsCatalog.GetTechnoInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetTechnoInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr2 = new Bitmap(img);

                if (_layer != null)
                {
                    if (_backgr != null)
                    {
                        if (_backgr2 != null)
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            int width2 = _backgr2.Width;
                            int height2 = _backgr2.Height;
                            _newBitmap = new Bitmap(width, height);
                            _newBitmap2 = new Bitmap(width1, height1);
                            using (var canvas = Graphics.FromImage(_newBitmap2))
                            {
                                if (height2 >= height1)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr2.Width, _backgr2.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }

                                int width3 = _newBitmap2.Width;
                                int height3 = _newBitmap2.Height;
                                using (var canvas1 = Graphics.FromImage(_newBitmap))
                                {
                                    if (height3 >= height)
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                    else
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _newBitmap2.Width, _newBitmap2.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                        else
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            _newBitmap = new Bitmap(width, height);
                            using (var canvas = Graphics.FromImage(_newBitmap))
                            {
                                if (height1 >= height)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                    }
                    else
                        pcbxVisualConfig.Image = _layer;
                }
                else
                {
                    if (_backgr != null)
                        pcbxVisualConfig.Image = _backgr;
                }

                if (pcbxVisualConfig.Image == null)
                    pcbxVisualConfig.Cursor = Cursors.Default;
                else
                    pcbxVisualConfig.Cursor = Cursors.Hand;
            }

            while (SplashWindow.bSmallCreated)
                SmallWaitForm.CloseS = true;
        }

        private void openFileDialog6_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var fileInfo = new System.IO.FileInfo(openFileDialog6.FileName);

            _attachmentsDt.Clear();

            DataRow newRow = _attachmentsDt.NewRow();
            newRow["FileName"] = Path.GetFileNameWithoutExtension(openFileDialog6.FileName);
            newRow["Extension"] = fileInfo.Extension;
            newRow["Path"] = openFileDialog6.FileName;
            _attachmentsDt.Rows.Add(newRow);
            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            int configId = _frontsCatalog.SaveTechnoInsetColorsConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
            if (configId != -1)
                _frontsCatalog.AttachTechnoInsetColorImage(_attachmentsDt, configId);

            pcbxVisualConfig.Image = null;
            if (cbVisualConfig.Checked)
            {
                Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
                T.Start();

                while (!SplashWindow.bSmallCreated) ;

                _layer = null;
                _backgr = null;
                _backgr2 = null;
                _newBitmap = null;
                _newBitmap2 = null;
                configId = _frontsCatalog.GetInsetTypesConfig(frontName, colorId, technoColorId);
                Image img = _frontsCatalog.GetInsetTypeImage(configId);
                if (configId != -1 && img != null)
                    _layer = new Bitmap(img);
                configId = _frontsCatalog.GetInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr = new Bitmap(img);

                configId = _frontsCatalog.GetTechnoInsetColorConfig(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId);
                img = _frontsCatalog.GetTechnoInsetColorImage(configId);
                if (configId != -1 && img != null)
                    _backgr2 = new Bitmap(img);

                if (_layer != null)
                {
                    if (_backgr != null)
                    {
                        if (_backgr2 != null)
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            int width2 = _backgr2.Width;
                            int height2 = _backgr2.Height;
                            _newBitmap = new Bitmap(width, height);
                            _newBitmap2 = new Bitmap(width1, height1);
                            using (var canvas = Graphics.FromImage(_newBitmap2))
                            {
                                if (height2 >= height1)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr2, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr2.Width, _backgr2.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width1, height1), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }

                                int width3 = _newBitmap2.Width;
                                int height3 = _newBitmap2.Height;
                                using (var canvas1 = Graphics.FromImage(_newBitmap))
                                {
                                    if (height3 >= height)
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                    else
                                    {
                                        canvas1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        canvas1.DrawImage(_newBitmap2, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _newBitmap2.Width, _newBitmap2.Height), GraphicsUnit.Pixel);
                                        canvas1.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                        canvas1.Save();
                                    }
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                        else
                        {
                            int width = _layer.Width;
                            int height = _layer.Height;
                            int width1 = _backgr.Width;
                            int height1 = _backgr.Height;
                            _newBitmap = new Bitmap(width, height);
                            using (var canvas = Graphics.FromImage(_newBitmap))
                            {
                                if (height1 >= height)
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                                else
                                {
                                    canvas.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    canvas.DrawImage(_backgr, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _backgr.Width, _backgr.Height), GraphicsUnit.Pixel);
                                    canvas.DrawImage(_layer, new Rectangle(0, 0, width, height), new Rectangle(0, 0, _layer.Width, _layer.Height), GraphicsUnit.Pixel);
                                    canvas.Save();
                                }
                            }

                            try
                            {
                                pcbxVisualConfig.Image = _newBitmap;
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                        }
                    }
                    else
                        pcbxVisualConfig.Image = _layer;
                }
                else
                {
                    if (_backgr != null)
                        pcbxVisualConfig.Image = _backgr;
                }

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;

                if (pcbxVisualConfig.Image == null)
                    pcbxVisualConfig.Cursor = Cursors.Default;
                else
                    pcbxVisualConfig.Cursor = Cursors.Hand;
            }
        }

        private void kryptonContextMenuItem20_Click(object sender, EventArgs e)
        {
            bool pressOk = false;

            string labelsCount = string.Empty;
            string positionsCount = string.Empty;
            string labelsLength = string.Empty;
            string serviceman = string.Empty;
            string milling = string.Empty;
            string docDateTime = string.Empty;
            string batch = string.Empty;
            string pallet = string.Empty;
            string profile = string.Empty;

            string labelsCount1 = string.Empty;
            string positionsCount1 = string.Empty;
            string labelsLength1 = string.Empty;
            string serviceman1 = string.Empty;
            string milling1 = string.Empty;
            string docDateTime1 = string.Empty;
            string batch1 = string.Empty;
            string pallet1 = string.Empty;
            string profile1 = string.Empty;

            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            SamplesLabelsForm sampleLabelInfo = new SamplesLabelsForm(this);
            _topForm = sampleLabelInfo;
            sampleLabelInfo.ShowDialog();
            pressOk = sampleLabelInfo.PressOK;

            docDateTime = sampleLabelInfo.DocDateTime;
            batch = sampleLabelInfo.Batch;
            pallet = sampleLabelInfo.Pallet;
            profile = sampleLabelInfo.Profile;
            serviceman = sampleLabelInfo.Serviceman;
            milling = sampleLabelInfo.Milling;
            labelsCount = sampleLabelInfo.LabelsCount;

            docDateTime1 = sampleLabelInfo.DocDateTime1;
            batch1 = sampleLabelInfo.Batch1;
            pallet1 = sampleLabelInfo.Pallet1;
            profile1 = sampleLabelInfo.Profile1;
            serviceman1 = sampleLabelInfo.Serviceman1;
            milling1 = sampleLabelInfo.Milling1;
            labelsCount1 = sampleLabelInfo.LabelsCount1;

            phantomForm.Close();
            phantomForm.Dispose();
            sampleLabelInfo.Dispose();
            _topForm = null;

            if (!pressOk)
                return;
            SamplesLabels samplesLabelsManager = new SamplesLabels();
            samplesLabelsManager.ClearLabelInfo();

            LabelContent labelInfo = new LabelContent()
            {
                DocDateTime = docDateTime,
                Batch = batch,
                Pallet = pallet,
                Profile = profile,
                Serviceman = serviceman,
                Milling = milling,

                DocDateTime1 = docDateTime1,
                Batch1 = batch1,
                Pallet1 = pallet1,
                Profile1 = profile1,
                Serviceman1 = serviceman1,
                Milling1 = milling1
            };
            samplesLabelsManager.AddLabelInfo(ref labelInfo);

            PrintDialog.Document = samplesLabelsManager.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                samplesLabelsManager.Print();
            }
        }

        private void btnSetDescription_Click(object sender, EventArgs e)
        {
            bool pressOk = false;

            string frontName = "";
            int colorId = -1;
            int technoColorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;
            int technoInsetTypeId = -1;
            int technoInsetColorId = -1;

            if (FrontsDataGrid.SelectedRows.Count > 0 && FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value != DBNull.Value)
                frontName = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].Value.ToString();
            if (FrameColorsDataGrid.SelectedRows.Count > 0 && FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                colorId = Convert.ToInt32(FrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (TechnoFrameColorsDataGrid.SelectedRows.Count > 0 && TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value != DBNull.Value)
                technoColorId = Convert.ToInt32(TechnoFrameColorsDataGrid.SelectedRows[0].Cells["ColorID"].Value);
            if (PatinaDataGrid.SelectedRows.Count > 0 && PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value != DBNull.Value)
                patinaId = Convert.ToInt32(PatinaDataGrid.SelectedRows[0].Cells["PatinaID"].Value);
            if (InsetTypesDataGrid.SelectedRows.Count > 0 && InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                insetTypeId = Convert.ToInt32(InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (InsetColorsDataGrid.SelectedRows.Count > 0 && InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                insetColorId = Convert.ToInt32(InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);
            if (TechnoInsetTypesDataGrid.SelectedRows.Count > 0 && TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value != DBNull.Value)
                technoInsetTypeId = Convert.ToInt32(TechnoInsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeID"].Value);
            if (TechnoInsetColorsDataGrid.SelectedRows.Count > 0 && TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value != DBNull.Value)
                technoInsetColorId = Convert.ToInt32(TechnoInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorID"].Value);

            bool isConfigImageToSite = false;
            bool bLatest = false;
            string category = string.Empty;
            string nameProd = string.Empty;
            string description = string.Empty;
            string sizes = string.Empty;
            string material = string.Empty;

            int configId = _frontsCatalog.GetFrontAttachments(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);
            if (configId != -1)
                isConfigImageToSite = _frontsCatalog.IsConfigImageToSite(configId, ref bLatest, ref category, ref nameProd, ref description, ref sizes, ref material);

            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            ProductDescriptionForm productDescriptionForm = new ProductDescriptionForm(this, isConfigImageToSite, bLatest, category, nameProd, description, sizes, material);
            _topForm = productDescriptionForm;
            productDescriptionForm.ShowDialog();

            pressOk = productDescriptionForm.PressOK;
            isConfigImageToSite = productDescriptionForm.ToSite;
            category = productDescriptionForm.Category;
            nameProd = productDescriptionForm.NameProd;
            description = productDescriptionForm.Description;
            sizes = productDescriptionForm.Sizes;
            material = productDescriptionForm.Material;
            bLatest = productDescriptionForm.Latest;

            phantomForm.Close();
            phantomForm.Dispose();
            productDescriptionForm.Dispose();
            _topForm = null;

            if (!pressOk)
                return;

            Thread T = new Thread(delegate () { SplashWindow.CreateSmallSplash(ref _topForm, "Загрузка данных с сервера.\r\nПодождите..."); });
            T.Start();

            while (!SplashWindow.bSmallCreated) ;

            if (configId != -1)
            {

                _frontsCatalog.ConfigImageToSite(configId, isConfigImageToSite, bLatest, category, nameProd, description, sizes, material);

                while (SplashWindow.bSmallCreated)
                    SmallWaitForm.CloseS = true;
            }
            else
            {
                int frontId = _frontsCatalog.GetFrontID(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);

                string front = FrontsDataGrid.SelectedRows[0].Cells["FrontName"].FormattedValue.ToString();
                string frameColor = FrameColorsDataGrid.SelectedRows[0].Cells["ColorName"].FormattedValue.ToString();
                string patina = PatinaDataGrid.SelectedRows[0].Cells["PatinaName"].FormattedValue.ToString();
                string insetType = InsetTypesDataGrid.SelectedRows[0].Cells["InsetTypeName"].FormattedValue.ToString();
                string insetColor = InsetColorsDataGrid.SelectedRows[0].Cells["InsetColorName"].FormattedValue.ToString();

                if (isConfigImageToSite && !_frontsCatalog.CreateFotoFromVisualConfig(frontId, colorId, technoColorId, patinaId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId,
                    category, front, frameColor, patina, insetType, insetColor))
                {
                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                    MessageBox.Show("Изображение не найдено!");
                }
                else
                {

                    configId = _frontsCatalog.GetFrontAttachments(frontName, colorId, technoColorId, insetTypeId, insetColorId, technoInsetTypeId, technoInsetColorId, patinaId);
                    _frontsCatalog.ConfigImageToSite(configId, isConfigImageToSite, bLatest, category, nameProd, description, sizes, material);

                    while (SplashWindow.bSmallCreated)
                        SmallWaitForm.CloseS = true;
                }
            }
        }

        private void KryptonButton1_Click(object sender, EventArgs e)
        {
            bool pressOk = false;

            int productId = -1;
            string decorName = string.Empty;
            int colorId = -1;
            int patinaId = -1;
            int insetTypeId = -1;
            int insetColorId = -1;

            if ((DataRowView)_decorCatalog.DecorProductsBindingSource.Current != null)
                productId = Convert.ToInt32(((DataRowView)_decorCatalog.DecorProductsBindingSource.Current)["ProductID"]);
            if ((DataRowView)_decorCatalog.DecorItemBindingSource.Current != null)
                decorName = ((DataRowView)_decorCatalog.DecorItemBindingSource.Current)["Name"].ToString();
            if ((DataRowView)_decorCatalog.ItemColorsBindingSource.Current != null)
                colorId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemColorsBindingSource.Current)["ColorID"]);
            if ((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current != null)
                patinaId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemPatinaBindingSource.Current)["PatinaID"]);
            if ((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current != null)
                insetTypeId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetTypesBindingSource.Current)["InsetTypeID"]);
            if ((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current != null)
                insetColorId = Convert.ToInt32(((DataRowView)_decorCatalog.ItemInsetColorsBindingSource.Current)["InsetColorID"]);

            string category = string.Empty;
            string description = string.Empty;
            string nameProd = string.Empty;
            string sizes = string.Empty;
            string material = string.Empty;
            bool isConfigImageToSite = false;
            bool bLatest = false;

            int configId = _decorCatalog.GetDecorAttachments(productId, decorName, colorId, patinaId, insetTypeId, insetColorId);
            if (configId != -1)
                isConfigImageToSite = _decorCatalog.IsConfigImageToSite(configId, ref bLatest, ref category, ref nameProd, ref description, ref sizes, ref material);
            else
            {
                MessageBox.Show("Изображение не найдено!");
                return;
            }
            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            ProductDescriptionForm productDescriptionForm = new ProductDescriptionForm(this, isConfigImageToSite, bLatest, category, nameProd, description, sizes, material);
            _topForm = productDescriptionForm;
            productDescriptionForm.ShowDialog();

            pressOk = productDescriptionForm.PressOK;
            isConfigImageToSite = productDescriptionForm.ToSite;
            category = productDescriptionForm.Category;
            description = productDescriptionForm.Description;
            nameProd = productDescriptionForm.NameProd;
            sizes = productDescriptionForm.Sizes;
            material = productDescriptionForm.Material;
            bLatest = productDescriptionForm.Latest;

            phantomForm.Close();
            phantomForm.Dispose();
            productDescriptionForm.Dispose();
            _topForm = null;

            if (!pressOk)
                return;

            if (configId != -1)
            {
                _decorCatalog.ConfigImageToSite(configId, isConfigImageToSite, bLatest, category, nameProd, description, sizes, material);
            }

        }

        private void KryptonContextMenuItem21_Click(object sender, EventArgs e)
        {
            bool pressOk = false;
            int labelsCount = 0;
            int positionsCount = 0;
            string docDateTime = string.Empty;

            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            CabFurLabelInfoMenu cabFurLabelInfoMenu = new CabFurLabelInfoMenu(this, false, false, false);
            _topForm = cabFurLabelInfoMenu;
            cabFurLabelInfoMenu.ShowDialog();
            pressOk = cabFurLabelInfoMenu.PressOK;
            labelsCount = cabFurLabelInfoMenu.LabelsCount;
            docDateTime = cabFurLabelInfoMenu.DocDateTime;
            positionsCount = cabFurLabelInfoMenu.PositionsCount;
            phantomForm.Close();
            phantomForm.Dispose();
            cabFurLabelInfoMenu.Dispose();
            _topForm = null;

            if (!pressOk)
                return;

            int decorId = _decorCatalog.GetDecorIdByName(DecorDataGrid.SelectedRows[0].Cells["Name"].Value.ToString());
            int labelsLength = Convert.ToInt32(LengthDataGrid.SelectedRows[0].Cells["Length"].Value);
            int labelsHeight = Convert.ToInt32(HeightDataGrid.SelectedRows[0].Cells["Height"].Value);
            int labelsWidth = Convert.ToInt32(WidthDataGrid.SelectedRows[0].Cells["Width"].Value);
            string color = ColorsDataGrid.SelectedRows[0].Cells["ColorName"].Value.ToString();

            int decorConfigId = 0;
            _cabFurDt.Clear();
            DataRow newRow = _cabFurDt.NewRow();

            DecorCatalog.TechStoreGroupInfo techStoreGroupInfo = _decorCatalog.GetSubGroupInfo(decorId);

            newRow["TechStoreName"] = techStoreGroupInfo.TechStoreName;
            newRow["TechStoreSubGroupName"] = techStoreGroupInfo.TechStoreSubGroupName;
            newRow["SubGroupNotes"] = techStoreGroupInfo.SubGroupNotes;
            newRow["SubGroupNotes1"] = techStoreGroupInfo.SubGroupNotes1;
            newRow["SubGroupNotes2"] = techStoreGroupInfo.SubGroupNotes2;
            newRow["Color"] = color;
            newRow["Length"] = labelsLength;
            newRow["Height"] = labelsHeight;
            newRow["Width"] = labelsWidth;
            newRow["LabelsCount"] = labelsCount;
            newRow["PositionsCount"] = positionsCount;
            newRow["DecorConfigID"] = decorConfigId;

            _cabFurDt.Rows.Add(newRow);

            _cabFurLabelManager.ClearLabelInfo();

            //Проверка
            if (_cabFurDt.Rows.Count == 0)
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                    "Очередь для печати пуста",
                    "Ошибка печати");
                return;
            }

            for (int i = 0; i < _cabFurDt.Rows.Count; i++)
            {
                labelsCount = Convert.ToInt32(_cabFurDt.Rows[i]["LabelsCount"]);
                decorConfigId = Convert.ToInt32(_cabFurDt.Rows[i]["DecorConfigID"]);
                int sampleLabelId = 0;
                for (int j = 0; j < labelsCount; j++)
                {
                    CabFurInfo labelInfo = new CabFurInfo();

                    DataTable dt = _cabFurDt.Clone();

                    DataRow destRow = dt.NewRow();
                    foreach (DataColumn column in _cabFurDt.Columns)
                    {
                        if (column.ColumnName == "FactoryType")
                            continue;
                        destRow[column.ColumnName] = _cabFurDt.Rows[i][column.ColumnName];
                    }
                    labelInfo.TechStoreName = _cabFurDt.Rows[i]["TechStoreName"].ToString();
                    labelInfo.TechStoreSubGroupName = _cabFurDt.Rows[i]["TechStoreSubGroupName"].ToString();
                    labelInfo.SubGroupNotes = _cabFurDt.Rows[i]["SubGroupNotes"].ToString();
                    labelInfo.SubGroupNotes1 = _cabFurDt.Rows[i]["SubGroupNotes1"].ToString();
                    labelInfo.SubGroupNotes2 = _cabFurDt.Rows[i]["SubGroupNotes2"].ToString();
                    sampleLabelId = _cabFurLabelManager.SaveSampleLabel(decorConfigId, DateTime.Now, Security.CurrentUserID, 2);
                    labelInfo.BarcodeNumber = _cabFurLabelManager.GetBarcodeNumber(19, sampleLabelId);
                    int factoryId = 1;
                    if (kryptonCheckSet2.CheckedButton == TPSCheckButton)
                        factoryId = 2;
                    labelInfo.FactoryType = factoryId;
                    labelInfo.ProductType = 2;
                    labelInfo.DocDateTime = docDateTime;
                    dt.Rows.Add(destRow);
                    labelInfo.OrderData = dt;

                    _cabFurLabelManager.AddLabelInfo(ref labelInfo);
                }
            }
            PrintDialog.Document = _cabFurLabelManager.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                _cabFurLabelManager.Print();
            }
        }

        private void KryptonContextMenuItem22_Click(object sender, EventArgs e)
        {
            bool pressOk = false;
            int labelsCount = 0;
            int positionsCount = 0;
            string docDateTime = string.Empty;

            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            CabFurLabelInfoMenu cabFurLabelInfoMenu = new CabFurLabelInfoMenu(this, false, false, false);
            _topForm = cabFurLabelInfoMenu;
            cabFurLabelInfoMenu.ShowDialog();
            pressOk = cabFurLabelInfoMenu.PressOK;
            labelsCount = cabFurLabelInfoMenu.LabelsCount;
            docDateTime = cabFurLabelInfoMenu.DocDateTime;
            positionsCount = cabFurLabelInfoMenu.PositionsCount;
            phantomForm.Close();
            phantomForm.Dispose();
            cabFurLabelInfoMenu.Dispose();
            _topForm = null;

            if (!pressOk)
                return;

            int decorId = _decorCatalog.GetDecorIdByName(DecorDataGrid.SelectedRows[0].Cells["Name"].Value.ToString());
            int labelsLength = Convert.ToInt32(LengthDataGrid.SelectedRows[0].Cells["Length"].Value);
            int labelsHeight = Convert.ToInt32(HeightDataGrid.SelectedRows[0].Cells["Height"].Value);
            int labelsWidth = Convert.ToInt32(WidthDataGrid.SelectedRows[0].Cells["Width"].Value);
            string color = ColorsDataGrid.SelectedRows[0].Cells["ColorName"].Value.ToString();
            string color2 = DecorInsetColorsDataGrid.SelectedRows[0].Cells["InsetColorName"].Value.ToString();

            int decorConfigId = 0;
            _cabFurDt.Clear();
            DataRow newRow = _cabFurDt.NewRow();

            DecorCatalog.TechStoreGroupInfo techStoreGroupInfo = _decorCatalog.GetSubGroupInfo(decorId);

            newRow["TechStoreName"] = techStoreGroupInfo.TechStoreName;
            newRow["TechStoreSubGroupName"] = techStoreGroupInfo.TechStoreSubGroupName;
            newRow["SubGroupNotes"] = techStoreGroupInfo.SubGroupNotes;
            newRow["SubGroupNotes1"] = techStoreGroupInfo.SubGroupNotes1;
            newRow["SubGroupNotes2"] = techStoreGroupInfo.SubGroupNotes2;
            newRow["Color"] = color;
            newRow["Color2"] = color2;
            newRow["Length"] = labelsLength;
            newRow["Height"] = labelsHeight;
            newRow["Width"] = labelsWidth;
            newRow["LabelsCount"] = labelsCount;
            newRow["PositionsCount"] = positionsCount;
            newRow["DecorConfigID"] = decorConfigId;

            _cabFurDt.Rows.Add(newRow);

            _cubeLabelManager.ClearLabelInfo();

            //Проверка
            if (_cabFurDt.Rows.Count == 0)
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                    "Очередь для печати пуста",
                    "Ошибка печати");
                return;
            }

            for (int i = 0; i < _cabFurDt.Rows.Count; i++)
            {
                labelsCount = Convert.ToInt32(_cabFurDt.Rows[i]["LabelsCount"]);
                decorConfigId = Convert.ToInt32(_cabFurDt.Rows[i]["DecorConfigID"]);
                int sampleLabelId = 0;
                for (int j = 0; j < labelsCount; j++)
                {
                    CabFurInfo labelInfo = new CabFurInfo();

                    DataTable dt = _cabFurDt.Clone();

                    DataRow destRow = dt.NewRow();
                    foreach (DataColumn column in _cabFurDt.Columns)
                    {
                        if (column.ColumnName == "FactoryType")
                            continue;
                        destRow[column.ColumnName] = _cabFurDt.Rows[i][column.ColumnName];
                    }
                    labelInfo.TechStoreName = _cabFurDt.Rows[i]["TechStoreName"].ToString();
                    labelInfo.TechStoreSubGroupName = _cabFurDt.Rows[i]["TechStoreSubGroupName"].ToString();
                    labelInfo.SubGroupNotes = _cabFurDt.Rows[i]["SubGroupNotes"].ToString();
                    labelInfo.SubGroupNotes1 = _cabFurDt.Rows[i]["SubGroupNotes1"].ToString();
                    labelInfo.SubGroupNotes2 = _cabFurDt.Rows[i]["SubGroupNotes2"].ToString();
                    sampleLabelId = _cubeLabelManager.SaveSampleLabel(decorConfigId, DateTime.Now, Security.CurrentUserID, 2);
                    labelInfo.BarcodeNumber = _cubeLabelManager.GetBarcodeNumber(19, sampleLabelId);
                    int factoryId = 1;
                    if (kryptonCheckSet2.CheckedButton == TPSCheckButton)
                        factoryId = 2;
                    labelInfo.FactoryType = factoryId;
                    labelInfo.ProductType = 2;
                    labelInfo.DocDateTime = docDateTime;
                    dt.Rows.Add(destRow);
                    labelInfo.OrderData = dt;

                    _cubeLabelManager.AddLabelInfo(ref labelInfo);
                }
            }
            PrintDialog.Document = _cubeLabelManager.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                _cubeLabelManager.Print();
            }
        }

        private void KryptonContextMenuItem23_Click(object sender, EventArgs e)
        {
            bool pressOk = false;
            int labelsCount = 0;
            int positionsCount = 0;
            string docDateTime = string.Empty;

            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            CabFurLabelInfoMenu cabFurLabelInfoMenu = new CabFurLabelInfoMenu(this, false, false, false);
            _topForm = cabFurLabelInfoMenu;
            cabFurLabelInfoMenu.ShowDialog();
            pressOk = cabFurLabelInfoMenu.PressOK;
            labelsCount = cabFurLabelInfoMenu.LabelsCount;
            docDateTime = cabFurLabelInfoMenu.DocDateTime;
            positionsCount = cabFurLabelInfoMenu.PositionsCount;
            phantomForm.Close();
            phantomForm.Dispose();
            cabFurLabelInfoMenu.Dispose();
            _topForm = null;

            if (!pressOk)
                return;

            int decorId = _decorCatalog.GetDecorIdByName(DecorDataGrid.SelectedRows[0].Cells["Name"].Value.ToString());
            int labelsLength = Convert.ToInt32(LengthDataGrid.SelectedRows[0].Cells["Length"].Value);
            int labelsHeight = Convert.ToInt32(HeightDataGrid.SelectedRows[0].Cells["Height"].Value);
            int labelsWidth = Convert.ToInt32(WidthDataGrid.SelectedRows[0].Cells["Width"].Value);
            string color = ColorsDataGrid.SelectedRows[0].Cells["ColorName"].Value.ToString();

            int decorConfigId = 0;
            _cabFurDt.Clear();
            DataRow newRow = _cabFurDt.NewRow();

            DecorCatalog.TechStoreGroupInfo techStoreGroupInfo = _decorCatalog.GetSubGroupInfo(decorId);

            newRow["TechStoreName"] = techStoreGroupInfo.TechStoreName;
            newRow["TechStoreSubGroupName"] = techStoreGroupInfo.TechStoreSubGroupName;
            newRow["SubGroupNotes"] = techStoreGroupInfo.SubGroupNotes;
            newRow["SubGroupNotes1"] = techStoreGroupInfo.SubGroupNotes1;
            newRow["SubGroupNotes2"] = techStoreGroupInfo.SubGroupNotes2;
            newRow["Color"] = color;
            newRow["Length"] = labelsLength;
            newRow["Height"] = labelsHeight;
            newRow["Width"] = labelsWidth;
            newRow["LabelsCount"] = labelsCount;
            newRow["PositionsCount"] = positionsCount;
            newRow["DecorConfigID"] = decorConfigId;

            _cabFurDt.Rows.Add(newRow);

            _cabFurLabelManager.ClearLabelInfo();

            //Проверка
            if (_cabFurDt.Rows.Count == 0)
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                    "Очередь для печати пуста",
                    "Ошибка печати");
                return;
            }

            for (int i = 0; i < _cabFurDt.Rows.Count; i++)
            {
                labelsCount = Convert.ToInt32(_cabFurDt.Rows[i]["LabelsCount"]);
                decorConfigId = Convert.ToInt32(_cabFurDt.Rows[i]["DecorConfigID"]);
                int sampleLabelId = 0;
                for (int j = 0; j < labelsCount; j++)
                {
                    CabFurInfo labelInfo = new CabFurInfo();

                    DataTable dt = _cabFurDt.Clone();

                    DataRow destRow = dt.NewRow();
                    foreach (DataColumn column in _cabFurDt.Columns)
                    {
                        if (column.ColumnName == "FactoryType")
                            continue;
                        destRow[column.ColumnName] = _cabFurDt.Rows[i][column.ColumnName];
                    }
                    labelInfo.TechStoreName = _cabFurDt.Rows[i]["TechStoreName"].ToString();
                    labelInfo.TechStoreSubGroupName = _cabFurDt.Rows[i]["TechStoreSubGroupName"].ToString();
                    labelInfo.SubGroupNotes = _cabFurDt.Rows[i]["SubGroupNotes"].ToString();
                    labelInfo.SubGroupNotes1 = _cabFurDt.Rows[i]["SubGroupNotes1"].ToString();
                    labelInfo.SubGroupNotes2 = _cabFurDt.Rows[i]["SubGroupNotes2"].ToString();
                    sampleLabelId = _cabFurLabelManager.SaveSampleLabel(decorConfigId, DateTime.Now, Security.CurrentUserID, 2);
                    labelInfo.BarcodeNumber = _cabFurLabelManager.GetBarcodeNumber(19, sampleLabelId);
                    int factoryId = 1;
                    if (kryptonCheckSet2.CheckedButton == TPSCheckButton)
                        factoryId = 2;
                    labelInfo.FactoryType = factoryId;
                    labelInfo.ProductType = 2;
                    labelInfo.DocDateTime = docDateTime;
                    dt.Rows.Add(destRow);
                    labelInfo.OrderData = dt;

                    _cabFurLabelManager.AddLabelInfo(ref labelInfo);
                }
            }
            PrintDialog.Document = _cabFurLabelManager.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                _cabFurLabelManager.Print();
            }
        }

        private void kryptonContextMenuItem24_Click(object sender, EventArgs e)
        {
            bool pressOk = false;
            int labelsCount = 0;
            int positionsCount = 0;
            string docDateTime = string.Empty;

            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            CabFurLabelInfoMenu cabFurLabelInfoMenu = new CabFurLabelInfoMenu(this, false, false, false);
            _topForm = cabFurLabelInfoMenu;
            cabFurLabelInfoMenu.ShowDialog();
            pressOk = cabFurLabelInfoMenu.PressOK;
            labelsCount = cabFurLabelInfoMenu.LabelsCount;
            docDateTime = cabFurLabelInfoMenu.DocDateTime;
            positionsCount = cabFurLabelInfoMenu.PositionsCount;
            phantomForm.Close();
            phantomForm.Dispose();
            cabFurLabelInfoMenu.Dispose();
            _topForm = null;

            if (!pressOk)
                return;

            int decorId = _decorCatalog.GetDecorIdByName(DecorDataGrid.SelectedRows[0].Cells["Name"].Value.ToString());
            int labelsLength = Convert.ToInt32(LengthDataGrid.SelectedRows[0].Cells["Length"].Value);
            int labelsHeight = Convert.ToInt32(HeightDataGrid.SelectedRows[0].Cells["Height"].Value);
            int labelsWidth = Convert.ToInt32(WidthDataGrid.SelectedRows[0].Cells["Width"].Value);
            string color = ColorsDataGrid.SelectedRows[0].Cells["ColorName"].Value.ToString();

            int decorConfigId = 0;
            _cabFurDt.Clear();
            DataRow newRow = _cabFurDt.NewRow();

            DecorCatalog.TechStoreGroupInfo techStoreGroupInfo = _decorCatalog.GetSubGroupInfo(decorId);

            newRow["TechStoreName"] = techStoreGroupInfo.TechStoreName;
            newRow["TechStoreSubGroupName"] = techStoreGroupInfo.TechStoreSubGroupName;
            newRow["SubGroupNotes"] = techStoreGroupInfo.SubGroupNotes;
            newRow["SubGroupNotes1"] = techStoreGroupInfo.SubGroupNotes1;
            newRow["SubGroupNotes2"] = techStoreGroupInfo.SubGroupNotes2;
            newRow["Color"] = color;
            newRow["Length"] = labelsLength;
            newRow["Height"] = labelsHeight;
            newRow["Width"] = labelsWidth;
            newRow["LabelsCount"] = labelsCount;
            newRow["PositionsCount"] = positionsCount;
            newRow["DecorConfigID"] = decorConfigId;

            _cabFurDt.Rows.Add(newRow);

            _cabFurLabelManager.ClearLabelInfo();

            //Проверка
            if (_cabFurDt.Rows.Count == 0)
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                    "Очередь для печати пуста",
                    "Ошибка печати");
                return;
            }

            for (int i = 0; i < _cabFurDt.Rows.Count; i++)
            {
                labelsCount = Convert.ToInt32(_cabFurDt.Rows[i]["LabelsCount"]);
                decorConfigId = Convert.ToInt32(_cabFurDt.Rows[i]["DecorConfigID"]);
                int sampleLabelId = 0;
                for (int j = 0; j < labelsCount; j++)
                {
                    CabFurInfo labelInfo = new CabFurInfo();

                    DataTable dt = _cabFurDt.Clone();

                    DataRow destRow = dt.NewRow();
                    foreach (DataColumn column in _cabFurDt.Columns)
                    {
                        if (column.ColumnName == "FactoryType")
                            continue;
                        destRow[column.ColumnName] = _cabFurDt.Rows[i][column.ColumnName];
                    }
                    labelInfo.TechStoreName = _cabFurDt.Rows[i]["TechStoreName"].ToString();
                    labelInfo.TechStoreSubGroupName = _cabFurDt.Rows[i]["TechStoreSubGroupName"].ToString();
                    labelInfo.SubGroupNotes = _cabFurDt.Rows[i]["SubGroupNotes"].ToString();
                    labelInfo.SubGroupNotes1 = _cabFurDt.Rows[i]["SubGroupNotes1"].ToString();
                    labelInfo.SubGroupNotes2 = _cabFurDt.Rows[i]["SubGroupNotes2"].ToString();
                    sampleLabelId = _cabFurLabelManager.SaveSampleLabel(decorConfigId, DateTime.Now, Security.CurrentUserID, 2);
                    labelInfo.BarcodeNumber = _cabFurLabelManager.GetBarcodeNumber(19, sampleLabelId);
                    int factoryId = 1;
                    if (kryptonCheckSet2.CheckedButton == TPSCheckButton)
                        factoryId = 2;
                    labelInfo.FactoryType = factoryId;
                    labelInfo.ProductType = 2;
                    labelInfo.DocDateTime = docDateTime;
                    dt.Rows.Add(destRow);
                    labelInfo.OrderData = dt;

                    _cabFurLabelManager.AddLabelInfo(ref labelInfo);
                }
            }
            PrintDialog.Document = _cabFurLabelManager.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                _cabFurLabelManager.Print();
            }
        }

        private void kryptonContextMenuItem25_Click(object sender, EventArgs e)
        {
            bool pressOk = false;
            int labelsCount = 0;
            int positionsCount = 0;
            string docDateTime = string.Empty;

            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            CabFurLabelInfoMenu cabFurLabelInfoMenu = new CabFurLabelInfoMenu(this, false, false, false);
            _topForm = cabFurLabelInfoMenu;
            cabFurLabelInfoMenu.ShowDialog();
            pressOk = cabFurLabelInfoMenu.PressOK;
            labelsCount = cabFurLabelInfoMenu.LabelsCount;
            docDateTime = cabFurLabelInfoMenu.DocDateTime;
            positionsCount = cabFurLabelInfoMenu.PositionsCount;
            phantomForm.Close();
            phantomForm.Dispose();
            cabFurLabelInfoMenu.Dispose();
            _topForm = null;

            if (!pressOk)
                return;

            int decorId = _decorCatalog.GetDecorIdByName(DecorDataGrid.SelectedRows[0].Cells["Name"].Value.ToString());
            int labelsLength = Convert.ToInt32(LengthDataGrid.SelectedRows[0].Cells["Length"].Value);
            int labelsHeight = Convert.ToInt32(HeightDataGrid.SelectedRows[0].Cells["Height"].Value);
            int labelsWidth = Convert.ToInt32(WidthDataGrid.SelectedRows[0].Cells["Width"].Value);
            string color = ColorsDataGrid.SelectedRows[0].Cells["ColorName"].Value.ToString();

            int decorConfigId = 0;
            _cabFurDt.Clear();
            DataRow newRow = _cabFurDt.NewRow();

            DecorCatalog.TechStoreGroupInfo techStoreGroupInfo = _decorCatalog.GetSubGroupInfo(decorId);

            newRow["TechStoreName"] = techStoreGroupInfo.TechStoreName;
            newRow["TechStoreSubGroupName"] = techStoreGroupInfo.TechStoreSubGroupName;
            newRow["SubGroupNotes"] = techStoreGroupInfo.SubGroupNotes;
            newRow["SubGroupNotes1"] = techStoreGroupInfo.SubGroupNotes1;
            newRow["SubGroupNotes2"] = techStoreGroupInfo.SubGroupNotes2;
            newRow["Color"] = color;
            newRow["Length"] = labelsLength;
            newRow["Height"] = labelsHeight;
            newRow["Width"] = labelsWidth;
            newRow["LabelsCount"] = labelsCount;
            newRow["PositionsCount"] = positionsCount;
            newRow["DecorConfigID"] = decorConfigId;

            _cabFurDt.Rows.Add(newRow);

            _cabFurLabelManager.ClearLabelInfo();

            //Проверка
            if (_cabFurDt.Rows.Count == 0)
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                    "Очередь для печати пуста",
                    "Ошибка печати");
                return;
            }

            for (int i = 0; i < _cabFurDt.Rows.Count; i++)
            {
                labelsCount = Convert.ToInt32(_cabFurDt.Rows[i]["LabelsCount"]);
                decorConfigId = Convert.ToInt32(_cabFurDt.Rows[i]["DecorConfigID"]);
                int sampleLabelId = 0;
                for (int j = 0; j < labelsCount; j++)
                {
                    CabFurInfo labelInfo = new CabFurInfo();

                    DataTable dt = _cabFurDt.Clone();

                    DataRow destRow = dt.NewRow();
                    foreach (DataColumn column in _cabFurDt.Columns)
                    {
                        if (column.ColumnName == "FactoryType")
                            continue;
                        destRow[column.ColumnName] = _cabFurDt.Rows[i][column.ColumnName];
                    }
                    labelInfo.TechStoreName = _cabFurDt.Rows[i]["TechStoreName"].ToString();
                    labelInfo.TechStoreSubGroupName = _cabFurDt.Rows[i]["TechStoreSubGroupName"].ToString();
                    labelInfo.SubGroupNotes = _cabFurDt.Rows[i]["SubGroupNotes"].ToString();
                    labelInfo.SubGroupNotes1 = _cabFurDt.Rows[i]["SubGroupNotes1"].ToString();
                    labelInfo.SubGroupNotes2 = _cabFurDt.Rows[i]["SubGroupNotes2"].ToString();
                    sampleLabelId = _cabFurLabelManager.SaveSampleLabel(decorConfigId, DateTime.Now, Security.CurrentUserID, 2);
                    labelInfo.BarcodeNumber = _cabFurLabelManager.GetBarcodeNumber(19, sampleLabelId);
                    int factoryId = 1;
                    if (kryptonCheckSet2.CheckedButton == TPSCheckButton)
                        factoryId = 2;
                    labelInfo.FactoryType = factoryId;
                    labelInfo.ProductType = 2;
                    labelInfo.DocDateTime = docDateTime;
                    dt.Rows.Add(destRow);
                    labelInfo.OrderData = dt;

                    _cabFurLabelManager.AddLabelInfo(ref labelInfo);
                }
            }
            PrintDialog.Document = _cabFurLabelManager.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                _cabFurLabelManager.Print();
            }
        }

        private void kryptonContextMenuItem26_Click(object sender, EventArgs e)
        {
            bool pressOk = false;
            int labelsCount = 0;
            int positionsCount = 0;
            string docDateTime = string.Empty;

            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            CabFurLabelInfoMenu cabFurLabelInfoMenu = new CabFurLabelInfoMenu(this, false, false, false);
            _topForm = cabFurLabelInfoMenu;
            cabFurLabelInfoMenu.ShowDialog();
            pressOk = cabFurLabelInfoMenu.PressOK;
            labelsCount = cabFurLabelInfoMenu.LabelsCount;
            docDateTime = cabFurLabelInfoMenu.DocDateTime;
            positionsCount = cabFurLabelInfoMenu.PositionsCount;
            phantomForm.Close();
            phantomForm.Dispose();
            cabFurLabelInfoMenu.Dispose();
            _topForm = null;

            if (!pressOk)
                return;

            int decorId = _decorCatalog.GetDecorIdByName(DecorDataGrid.SelectedRows[0].Cells["Name"].Value.ToString());
            int labelsLength = Convert.ToInt32(LengthDataGrid.SelectedRows[0].Cells["Length"].Value);
            int labelsHeight = Convert.ToInt32(HeightDataGrid.SelectedRows[0].Cells["Height"].Value);
            int labelsWidth = Convert.ToInt32(WidthDataGrid.SelectedRows[0].Cells["Width"].Value);
            string color = ColorsDataGrid.SelectedRows[0].Cells["ColorName"].Value.ToString();

            int decorConfigId = 0;
            _cabFurDt.Clear();
            DataRow newRow = _cabFurDt.NewRow();

            DecorCatalog.TechStoreGroupInfo techStoreGroupInfo = _decorCatalog.GetSubGroupInfo(decorId);

            newRow["TechStoreName"] = techStoreGroupInfo.TechStoreName;
            newRow["TechStoreSubGroupName"] = techStoreGroupInfo.TechStoreSubGroupName;
            newRow["SubGroupNotes"] = techStoreGroupInfo.SubGroupNotes;
            newRow["SubGroupNotes1"] = techStoreGroupInfo.SubGroupNotes1;
            newRow["SubGroupNotes2"] = techStoreGroupInfo.SubGroupNotes2;
            newRow["Color"] = color;
            newRow["Length"] = labelsLength;
            newRow["Height"] = labelsHeight;
            newRow["Width"] = labelsWidth;
            newRow["LabelsCount"] = labelsCount;
            newRow["PositionsCount"] = positionsCount;
            newRow["DecorConfigID"] = decorConfigId;

            _cabFurDt.Rows.Add(newRow);

            _cabFurLabelManager.ClearLabelInfo();

            //Проверка
            if (_cabFurDt.Rows.Count == 0)
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                    "Очередь для печати пуста",
                    "Ошибка печати");
                return;
            }

            for (int i = 0; i < _cabFurDt.Rows.Count; i++)
            {
                labelsCount = Convert.ToInt32(_cabFurDt.Rows[i]["LabelsCount"]);
                decorConfigId = Convert.ToInt32(_cabFurDt.Rows[i]["DecorConfigID"]);
                int sampleLabelId = 0;
                for (int j = 0; j < labelsCount; j++)
                {
                    CabFurInfo labelInfo = new CabFurInfo();

                    DataTable dt = _cabFurDt.Clone();

                    DataRow destRow = dt.NewRow();
                    foreach (DataColumn column in _cabFurDt.Columns)
                    {
                        if (column.ColumnName == "FactoryType")
                            continue;
                        destRow[column.ColumnName] = _cabFurDt.Rows[i][column.ColumnName];
                    }


                    labelInfo.TechStoreName = _cabFurDt.Rows[i]["TechStoreName"].ToString();
                    labelInfo.TechStoreSubGroupName = _cabFurDt.Rows[i]["TechStoreSubGroupName"].ToString();
                    labelInfo.SubGroupNotes = _cabFurDt.Rows[i]["SubGroupNotes"].ToString();
                    labelInfo.SubGroupNotes1 = _cabFurDt.Rows[i]["SubGroupNotes1"].ToString();
                    labelInfo.SubGroupNotes2 = _cabFurDt.Rows[i]["SubGroupNotes2"].ToString();
                    sampleLabelId = _cabFurLabelManager.SaveSampleLabel(decorConfigId, DateTime.Now, Security.CurrentUserID, 2);
                    labelInfo.BarcodeNumber = _cabFurLabelManager.GetBarcodeNumber(19, sampleLabelId);

                    int factoryId = 1;
                    if (kryptonCheckSet2.CheckedButton == TPSCheckButton)
                        factoryId = 2;
                    labelInfo.FactoryType = factoryId;
                    labelInfo.ProductType = 2;
                    labelInfo.DocDateTime = docDateTime;
                    dt.Rows.Add(destRow);
                    labelInfo.OrderData = dt;

                    _cabFurLabelManager.AddLabelInfo(ref labelInfo);
                }
            }
            PrintDialog.Document = _cabFurLabelManager.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                _cabFurLabelManager.Print();
            }
        }

        private void kryptonContextMenuItem27_Click(object sender, EventArgs e)
        {
            bool pressOk = false;
            int labelsCount = 0;
            int positionsCount = 0;
            string docDateTime = string.Empty;

            PhantomForm phantomForm = new Infinium.PhantomForm();
            phantomForm.Show();
            CabFurLabelInfoMenu cabFurLabelInfoMenu = new CabFurLabelInfoMenu(this, false, false, false);
            _topForm = cabFurLabelInfoMenu;
            cabFurLabelInfoMenu.ShowDialog();
            pressOk = cabFurLabelInfoMenu.PressOK;
            labelsCount = cabFurLabelInfoMenu.LabelsCount;
            docDateTime = cabFurLabelInfoMenu.DocDateTime;
            positionsCount = cabFurLabelInfoMenu.PositionsCount;
            phantomForm.Close();
            phantomForm.Dispose();
            cabFurLabelInfoMenu.Dispose();
            _topForm = null;

            if (!pressOk)
                return;

            int decorId = _decorCatalog.GetDecorIdByName(DecorDataGrid.SelectedRows[0].Cells["Name"].Value.ToString());
            int labelsLength = Convert.ToInt32(LengthDataGrid.SelectedRows[0].Cells["Length"].Value);
            int labelsHeight = Convert.ToInt32(HeightDataGrid.SelectedRows[0].Cells["Height"].Value);
            int labelsWidth = Convert.ToInt32(WidthDataGrid.SelectedRows[0].Cells["Width"].Value);
            string color = ColorsDataGrid.SelectedRows[0].Cells["ColorName"].Value.ToString();

            int decorConfigId = 0;
            _cabFurDt.Clear();
            DataRow newRow = _cabFurDt.NewRow();

            DecorCatalog.TechStoreGroupInfo techStoreGroupInfo = _decorCatalog.GetSubGroupInfo(decorId);

            newRow["TechStoreName"] = techStoreGroupInfo.TechStoreName;
            newRow["TechStoreSubGroupName"] = techStoreGroupInfo.TechStoreSubGroupName;
            newRow["SubGroupNotes"] = techStoreGroupInfo.SubGroupNotes;
            newRow["SubGroupNotes1"] = techStoreGroupInfo.SubGroupNotes1;
            newRow["SubGroupNotes2"] = techStoreGroupInfo.SubGroupNotes2;
            newRow["Color"] = color;
            newRow["Length"] = labelsLength;
            newRow["Height"] = labelsHeight;
            newRow["Width"] = labelsWidth;
            newRow["LabelsCount"] = labelsCount;
            newRow["PositionsCount"] = positionsCount;
            newRow["DecorConfigID"] = decorConfigId;

            _cabFurDt.Rows.Add(newRow);

            _cabFurLabelManager.ClearLabelInfo();

            //Проверка
            if (_cabFurDt.Rows.Count == 0)
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                    "Очередь для печати пуста",
                    "Ошибка печати");
                return;
            }

            for (int i = 0; i < _cabFurDt.Rows.Count; i++)
            {
                labelsCount = Convert.ToInt32(_cabFurDt.Rows[i]["LabelsCount"]);
                decorConfigId = Convert.ToInt32(_cabFurDt.Rows[i]["DecorConfigID"]);
                int sampleLabelId = 0;
                for (int j = 0; j < labelsCount; j++)
                {
                    CabFurInfo labelInfo = new CabFurInfo();

                    DataTable dt = _cabFurDt.Clone();

                    DataRow destRow = dt.NewRow();
                    foreach (DataColumn column in _cabFurDt.Columns)
                    {
                        if (column.ColumnName == "FactoryType")
                            continue;
                        destRow[column.ColumnName] = _cabFurDt.Rows[i][column.ColumnName];
                    }


                    labelInfo.TechStoreName = _cabFurDt.Rows[i]["TechStoreName"].ToString();
                    labelInfo.TechStoreSubGroupName = _cabFurDt.Rows[i]["TechStoreSubGroupName"].ToString();
                    labelInfo.SubGroupNotes = _cabFurDt.Rows[i]["SubGroupNotes"].ToString();
                    labelInfo.SubGroupNotes1 = _cabFurDt.Rows[i]["SubGroupNotes1"].ToString();
                    labelInfo.SubGroupNotes2 = _cabFurDt.Rows[i]["SubGroupNotes2"].ToString();
                    sampleLabelId = _cabFurLabelManager.SaveSampleLabel(decorConfigId, DateTime.Now, Security.CurrentUserID, 2);
                    labelInfo.BarcodeNumber = _cabFurLabelManager.GetBarcodeNumber(19, sampleLabelId);

                    int factoryId = 1;
                    if (kryptonCheckSet2.CheckedButton == TPSCheckButton)
                        factoryId = 2;
                    labelInfo.FactoryType = factoryId;
                    labelInfo.ProductType = 2;
                    labelInfo.DocDateTime = docDateTime;
                    dt.Rows.Add(destRow);
                    labelInfo.OrderData = dt;

                    _cabFurLabelManager.AddLabelInfo(ref labelInfo);
                }
            }
            PrintDialog.Document = _cabFurLabelManager.PD;

            if (PrintDialog.ShowDialog() == DialogResult.OK)
            {
                _cabFurLabelManager.Print();
            }
        }

        private void MenuButton_Click(object sender, EventArgs e)
        {
            MenuPanel.Visible = !MenuPanel.Visible;
        }

        private void cbClients_CheckedChanged(object sender, EventArgs e)
        {
            cbExcluzive.Enabled = !cbExcluzive.Checked;
            FilterClientsDataGrid.Enabled = !FilterClientsDataGrid.Enabled;
        }

        private void btnFilterCatalog_Click(object sender, EventArgs e)
        {
            int ClientID = -1;

            if (cbClients.Checked && FilterClientsDataGrid.SelectedRows.Count > 0)
                ClientID = Convert.ToInt32(FilterClientsDataGrid.SelectedRows[0].Cells["ClientID"].Value);

            if (cbClients.Checked)
                _frontsCatalog.FilterFronts(ClientID);
            else
                _frontsCatalog.FilterFronts();
        }
    }
}
