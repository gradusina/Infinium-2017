using System.Windows.Forms;
using Infinium;

namespace Testing
{
    public partial class LightStartForm : Form
    {
        private readonly Connection _connection;

        public LightStartForm()
        {
            InitializeComponent();

            _connection = new Connection();
            ConnectionStrings.UsersConnectionString = _connection.GetConnectionString("ConnectionUsers.config");
            ConnectionStrings.CatalogConnectionString = _connection.GetConnectionString("ConnectionCatalog.config");
            ConnectionStrings.LightConnectionString = _connection.GetConnectionString("ConnectionLight.config");
            ConnectionStrings.MarketingOrdersConnectionString =
                _connection.GetConnectionString("ConnectionMarketingOrders.config");
            ConnectionStrings.MarketingReferenceConnectionString =
                _connection.GetConnectionString("ConnectionMarketingReference.config");
            ConnectionStrings.StorageConnectionString = _connection.GetConnectionString("ConnectionStorage.config");
            ConnectionStrings.ZOVOrdersConnectionString = _connection.GetConnectionString("ConnectionZOVOrders.config");
            ConnectionStrings.ZOVReferenceConnectionString =
                _connection.GetConnectionString("ConnectionZOVReference.config");

            DatabaseConfigsManager DatabaseConfigsManager = new DatabaseConfigsManager();
            DatabaseConfigsManager.ReadAnimationFlag("Animation.config");
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            LoadCalculationsForm form = new LoadCalculationsForm(this);

            form.Show();
        }

    }
}
