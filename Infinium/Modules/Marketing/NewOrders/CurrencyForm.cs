
using Infinium.Modules.Marketing.NewOrders;
using Infinium.Modules.Marketing.NewOrders.InvoiceReportToDbf;

using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Infinium
{
    public partial class CurrencyForm : Form
    {
        private const int EHide = 2;
        private const int EShow = 1;
        private const int EClose = 3;
        private const int EMainMenu = 4;

        private readonly bool _admin;

        private bool _rateExist;
        private bool _fixedPaymentRate;
        private bool _cbrDailyRates = true;
        private bool _nbrbDailyRates = true;
        private decimal _eurrubCurrency = 1000000;
        private decimal _usdrubCurrency = 1000000;
        private decimal _eurusdCurrency = 1000000;
        private decimal _eurbyrCurrency = 1000000;
        private decimal BYN = 1000000;
        private decimal _currencyTotalCost;

        private int _formEvent;
        private readonly int _clientId;
        private int _transportType;
        private decimal _weight;
        private decimal _transportCost;
        private decimal _rate = 1;

        private readonly Form _mainForm;
        private Form _topForm;

        private readonly OrdersManager _ordersManager;
        private readonly OrdersCalculate _ordersCalculate;
        private readonly InvoiceReportToDbf _dbfReport;

        private readonly bool _canSetDirectorDiscount;

        public CurrencyForm(Form tMainForm, ref OrdersManager tOrdersManager, ref OrdersCalculate tOrdersCalculate, ref InvoiceReportToDbf tDbfReport,
            int iClientId, string clientName, bool bCanSetDirectorDiscount, bool bAdmin)
        {
            _mainForm = tMainForm;
            _clientId = iClientId;
            _canSetDirectorDiscount = bCanSetDirectorDiscount;
            _admin = bAdmin;
            _ordersManager = tOrdersManager;
            _ordersCalculate = tOrdersCalculate;
            _dbfReport = tDbfReport;

            InitializeComponent();
            panel8.Visible = false;
            if (bCanSetDirectorDiscount)
            {
                panel7.Visible = true;
            }
            if (_admin)
            {
                CurrencyDateTimePicker.Enabled = true;
            }
            label2.Text = "Расчет " + clientName;
            CurrencyTypeComboBox.Items.Add("Евро - Евро");
            CurrencyTypeComboBox.Items.Add("Евро - Доллар США");
            CurrencyTypeComboBox.Items.Add("Евро - Российский рубль");
            CurrencyTypeComboBox.Items.Add("Евро - Белорусский рубль");

            Initialize();
        }

        private void ConfirmationOKButton_Click(object sender, EventArgs e)
        {
            if (!_canSetDirectorDiscount && Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]) == 2)
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                        "Заказ согласован, пересчет запрещен. Звоните Директорату",
                        "Пересчет заказа");
                _formEvent = EClose;
                AnimateTimer.Enabled = true;
                return;
            }

            if (CurrencyTextBox.Text == "0")
            {
                MessageBox.Show("Курс валюты равен нулю, проверьте курс");
                return;
            }

            //if (rbCpt.Checked && Convert.ToDecimal(TransportCostTextEdit.Text) == 0)
            //{
            //    MessageBox.Show("Введите сумму транспорта");
            //    return;
            //}

            if (Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]) != 2)
                _ordersManager.SetNotConfirmOrder();

            _ordersManager.SetFixedPaymentRate(
                Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]),
                _fixedPaymentRate);

            int currencyTypeId = CurrencyTypeComboBox.SelectedIndex + 1;
            if (CurrencyTypeComboBox.SelectedIndex == 3)
                currencyTypeId = 5;
            int delayOfPayment = 0;
            int discountFactoringId = 0;
            //int discountPaymentConditionId = 1;
            int discountPaymentConditionId = 3;
            //if (rbHalfOfPayment.Checked)
            //{
            //    discountPaymentConditionId = 2;
            //    delayOfPayment = 0;
            //}
            //if (rbFullPayment.Checked)
            //{
            //    discountPaymentConditionId = 3;
            //    delayOfPayment = 0;
            //}
            //if (rbFullPayment.Checked)
            //{
            //    discountPaymentConditionId = 3;
            //    delayOfPayment = 0;
            //}
            //if (kryptonRadioButton3.Checked)
            //{
            //    discountPaymentConditionId = 5;
            //    delayOfPayment = 0;
            //}
            //if (kryptonRadioButton5.Checked)
            //{
            //    discountPaymentConditionId = 6;
            //    delayOfPayment = 0;
            //}
            //if (rbFactoring.Checked)
            //{
            //    discountPaymentConditionId = 4;
            //}

            //if (discountPaymentConditionId == 1)
            //    delayOfPayment = Convert.ToInt32(tbDelayOfPayment.Text);

            decimal profilTotalSum = 0;
            decimal tpsTotalSum = 0;
            _ordersCalculate.GetMainOrderTotalSum(Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), ref profilTotalSum, ref tpsTotalSum);

            if (profilTotalSum < Convert.ToDecimal(tbComplaintProfilCost.Text) || tpsTotalSum < Convert.ToDecimal(tbComplaintTPSCost.Text))
            {
                Infinium.LightMessageBox.Show(ref _topForm, false,
                       "Сумма рекламации не может быть больше стоимости заказа",
                       "Расчет заказа");
                return;
            }

            decimal discountPaymentCondition = _ordersManager.DiscountPaymentCondition(discountPaymentConditionId);
            string factoringName = string.Empty;
            if (discountPaymentConditionId == 4)
            {
                if (kryptonRadioButton2.Checked)
                    factoringName = kryptonRadioButton2.Text;
                if (kryptonRadioButton4.Checked)
                    factoringName = kryptonRadioButton4.Text;
                if (kryptonRadioButton1.Checked)
                    factoringName = kryptonRadioButton1.Text;
                int daysCount = 0;
                discountPaymentCondition = _ordersManager.DiscountFactoring(ref discountFactoringId, ref daysCount, currencyTypeId, factoringName);
                delayOfPayment = daysCount;
            }
            //decimal profilDiscountOrderSum = _ordersManager.DiscountOrderSum(profilTotalSum);
            //decimal tpsDiscountOrderSum = _ordersManager.DiscountOrderSum(tpsTotalSum);
            decimal profilDiscountOrderSum = 0;
            decimal tpsDiscountOrderSum = 0;
            decimal profilDiscountDirector = Convert.ToDecimal(tbProfilDiscountDirector.Text);
            decimal tpsDiscountDirector = Convert.ToDecimal(tbTPSDiscountDirector.Text);
            decimal profilTotalDiscount = 0;
            decimal tpsTotalDiscount = 0;
            int transportType = 0;
            object confirmDateTime = ((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["ConfirmDateTime"];


            if (discountPaymentConditionId == 1 || discountPaymentConditionId == 6)
            {
                profilDiscountOrderSum = 0;
                tpsDiscountOrderSum = 0;
            }
            decimal transportCost = Convert.ToDecimal(TransportCostTextEdit.Text);
            if (rbFCA.Checked)
            {
                transportType = 1;
                transportCost = 0;
            }
            profilTotalDiscount = discountPaymentCondition + profilDiscountOrderSum + profilDiscountDirector;
            tpsTotalDiscount = discountPaymentCondition + tpsDiscountOrderSum + tpsDiscountDirector;
            
            _ordersManager.CurrencyTypeID = currencyTypeId;

            if (confirmDateTime != DBNull.Value)
                _ordersCalculate.ConfirmDateTime = Convert.ToDateTime(confirmDateTime);

            Tuple<bool, decimal, decimal, decimal> clientRates =
                _ordersCalculate.GetFixedPaymentRate(_ordersManager.CurrentClientID, _ordersCalculate.ConfirmDateTime);

            Tuple<bool, decimal, decimal, decimal, decimal> dateRates =
                _ordersManager.GetDateRates(_ordersCalculate.ConfirmDateTime);

            bool fixedClientRate = clientRates.Item1;

            if (fixedClientRate)
            {
                switch (_ordersManager.CurrencyTypeID)
                {
                    case 1:
                        if (dateRates.Item1)
                            _ordersCalculate.Rate = clientRates.Item4 / dateRates.Item4;
                        break;
                    case 2:
                        _ordersCalculate.Rate = clientRates.Item2;
                        break;
                    case 3:
                        _ordersCalculate.Rate = clientRates.Item3;
                        break;
                    case 5:
                        _ordersCalculate.Rate = clientRates.Item4;
                        break;
                    default:
                        break;
                }
            }
            else
            {
                if (dateRates.Item1)
                {
                    switch (_ordersManager.CurrencyTypeID)
                    {
                        case 1:
                            _ordersCalculate.Rate = 1;
                            break;
                        case 2:
                            _ordersCalculate.Rate = dateRates.Item2;
                            break;
                        case 3:
                            _ordersCalculate.Rate = dateRates.Item3;
                            break;
                        case 5:
                            _ordersCalculate.Rate = dateRates.Item4;
                            break;
                        default:
                            break;
                    }
                }

            }

            decimal originalRate = _ordersCalculate.Rate;
            decimal paymentRate = _ordersCalculate.Rate;
            if (_clientId != 145 && discountPaymentConditionId == 1 && (currencyTypeId == 3 || currencyTypeId == 5))
            {
                paymentRate = originalRate * 1.05m;
            }

            _ordersCalculate.Recalculate(Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]),
                Convert.ToDecimal(tbComplaintProfilCost.Text), Convert.ToDecimal(tbComplaintTPSCost.Text),
                transportType, transportCost, Convert.ToDecimal(AdditionalCostTextEdit.Text),
                profilDiscountDirector, tpsDiscountDirector,
                (discountPaymentCondition + profilDiscountOrderSum), (discountPaymentCondition + tpsDiscountOrderSum), discountPaymentCondition,
                currencyTypeId, paymentRate, confirmDateTime);

            _currencyTotalCost = _dbfReport.CalcCurrencyCost(
                Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]),
                Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["ClientID"]), paymentRate);
            _ordersManager.SetCurrencyCost(Convert.ToDecimal(tbComplaintProfilCost.Text), Convert.ToDecimal(tbComplaintTPSCost.Text), tbComplaintNotes.Text, IsComplaintCheckBox.Checked, delayOfPayment,
                transportCost, Convert.ToDecimal(AdditionalCostTextEdit.Text),
                currencyTypeId, originalRate, paymentRate,
                discountPaymentConditionId, discountFactoringId, profilDiscountOrderSum, tpsDiscountOrderSum, profilDiscountDirector, tpsDiscountDirector, _currencyTotalCost);

            //change date
            string desireDate = "";

            if (NoDateCheckBox.Checked == false)
                desireDate = DispatchDateTimePicker.Value.Date.ToString("yyyy-MM-dd");

            _ordersManager.LastCalcDate(Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]));
            _ordersManager.ChangeDate(desireDate);

            if (_canSetDirectorDiscount || Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]) != 0)
            {
                _ordersManager.MoveOrdersTo(Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current)["MegaOrderID"]),
                    Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]));
            }

            _ordersManager.FixOrderEvent(Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["MegaOrderID"]), "Заказ рассчитан");
            _ordersManager.SendReport = true;

            _formEvent = EClose;
            AnimateTimer.Enabled = true;
        }

        private void ConfirmationCancelButton_Click(object sender, EventArgs e)
        {
            _ordersManager.SendReport = false;
            _formEvent = EClose;
            AnimateTimer.Enabled = true;
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
                        this.Close();
                    }

                    if (_formEvent == EHide)
                    {
                        _mainForm.Activate();
                        this.Hide();
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
                        this.Close();
                    }

                    if (_formEvent == EHide)
                    {
                        _mainForm.Activate();
                        this.Hide();
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

        public void SetParameters(int iTransportType, int delayOfPayment, object desireDate, object confirmDate, int currencyTypeId, decimal originalRate, decimal paymentRate, decimal complaintProfilCost, decimal complaintTpsCost, string complaintNotes,
            decimal dTransportCost, decimal additionalCost, int discountPaymentConditionId, int discountFactoringId, decimal profilDiscountDirector, decimal tpsDiscountDirector, decimal dWeight)
        {
            _transportType = iTransportType;
            _transportCost = dTransportCost;
            _weight = dWeight;
            if (iTransportType == 1)
            {
                TransportCostTextEdit.Visible = false;
                rbFCA.Checked = true;
            }

            switch (currencyTypeId)
            {
                case 1:
                    CurrencyTypeComboBox.SelectedIndex = 0;
                    break;
                case 2:
                    CurrencyTypeComboBox.SelectedIndex = 1;
                    break;
                case 3:
                    CurrencyTypeComboBox.SelectedIndex = 2;
                    break;
                case 5:
                    CurrencyTypeComboBox.SelectedIndex = 3;
                    break;
                default:
                    CurrencyTypeComboBox.SelectedIndex = 0;
                    break;
            }
            
            if (Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]) == 2)
            {
                if (!_canSetDirectorDiscount)
                {
                    CurrencyTypeComboBox.Enabled = false;
                    rbDelayOfPayment.Enabled = false;
                    rbFactoring.Enabled = false;
                    rbFullPayment.Enabled = false;
                    tbDelayOfPayment.Enabled = false;
                    kryptonRadioButton3.Enabled = false;
                    kryptonRadioButton5.Enabled = false;

                    kryptonRadioButton2.Enabled = false;
                    kryptonRadioButton4.Enabled = false;
                    kryptonRadioButton1.Enabled = false;
                    rbCpt.Enabled = false;
                    rbFCA.Enabled = false;
                    TransportCostTextEdit.Enabled = false;
                    AdditionalCostTextEdit.Enabled = false;
                    IsComplaintCheckBox.Enabled = false;
                    tbComplaintProfilCost.Enabled = false;
                    tbComplaintTPSCost.Enabled = false;
                    tbComplaintNotes.Enabled = false;

                    panel2.Enabled = false;
                    panel3.Enabled = false;
                    panel4.Enabled = false;
                    panel6.Enabled = false;
                    panel7.Enabled = false;
                }
            }
            else
            {
                if (_transportType == 0)
                {
                    decimal d = _ordersManager.TransportCost(dWeight);
                    if (Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]) == 0
                        || Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]) == 3)
                        _transportCost = _ordersManager.TransportCost(dWeight);
                    else
                        _transportCost = dTransportCost;
                }
                else
                    _transportCost = 0;
            }

            if (Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]) == 0)
                delayOfPayment = OrdersManager.GetDelayOfPayment(Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["ClientID"]));
            else
                delayOfPayment = Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["DelayOfPayment"]);


            tbDelayOfPayment.Text = delayOfPayment.ToString();
            TransportCostTextEdit.Text = _transportCost.ToString();
            AdditionalCostTextEdit.Text = additionalCost.ToString();
            tbComplaintProfilCost.Text = complaintProfilCost.ToString();
            tbComplaintTPSCost.Text = complaintTpsCost.ToString();
            tbComplaintNotes.Text = complaintNotes.ToString();
            tbProfilDiscountDirector.Text = profilDiscountDirector.ToString();
            tbTPSDiscountDirector.Text = tpsDiscountDirector.ToString();
            
            if (confirmDate != DBNull.Value)
                _ordersManager.GetDateRates(Convert.ToDateTime(confirmDate), ref _rateExist, ref _eurusdCurrency, ref _eurrubCurrency, ref _eurbyrCurrency, ref _usdrubCurrency);
            else
                _ordersManager.GetDateRates(CurrencyDateTimePicker.Value.Date, ref _rateExist, ref _eurusdCurrency, ref _eurrubCurrency, ref _eurbyrCurrency, ref _usdrubCurrency);

            BYN = _eurbyrCurrency;
            
            if (!_rateExist)
            {
                if (confirmDate != DBNull.Value)
                {
                    _cbrDailyRates = OrdersManager.CBRDailyRates(Convert.ToDateTime(confirmDate), ref _eurrubCurrency, ref _usdrubCurrency);

                    _eurbyrCurrency = CurrencyConverter.NbrbDailyRates(DateTime.Now);

                    //NBRBDailyRates = OrdersManager.NBRBDailyRates(Convert.ToDateTime(ConfirmDate), ref EURBYRCurrency);
                }
                else
                {
                    _cbrDailyRates = OrdersManager.CBRDailyRates(CurrencyDateTimePicker.Value.Date, ref _eurrubCurrency, ref _usdrubCurrency);
                    _eurbyrCurrency = CurrencyConverter.NbrbDailyRates(DateTime.Now);
                    //NBRBDailyRates = OrdersManager.NBRBDailyRates(CurrencyDateTimePicker.Value.Date, ref EURBYRCurrency);
                }
                if (_usdrubCurrency != 0)
                    _eurusdCurrency = Decimal.Round(_eurrubCurrency / _usdrubCurrency, 4, MidpointRounding.AwayFromZero);

                if (_eurbyrCurrency == 1000000)
                {

                    Infinium.LightMessageBox.Show(ref _topForm, false,
                        "Курс не взят. Возможная причина - нет связи с банком без авторизации в kerio control. Войдите в kerio и повторите попытку",
                        "Расчет заказа");
                    return;
                }

                OrdersManager.SaveDateRates(
                    confirmDate != DBNull.Value ? Convert.ToDateTime(confirmDate) : CurrencyDateTimePicker.Value.Date,
                    _eurusdCurrency, _eurrubCurrency, _eurbyrCurrency, _usdrubCurrency);
                _rateExist = true;
            }

            if (confirmDate != DBNull.Value)
            {
                _ordersCalculate.GetFixedPaymentRate(Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["ClientID"]),
                    Convert.ToDateTime(confirmDate), ref _fixedPaymentRate, ref _eurusdCurrency, ref _eurrubCurrency, ref _eurbyrCurrency);
                CurrencyDateTimePicker.Value = Convert.ToDateTime(confirmDate);
            }
            else
                _ordersCalculate.GetFixedPaymentRate(Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["ClientID"]),
                    CurrencyDateTimePicker.Value.Date, ref _fixedPaymentRate, ref _eurusdCurrency, ref _eurrubCurrency, ref _eurbyrCurrency);
            if (_fixedPaymentRate)
            {
                _usdrubCurrency = 1;
                NoConnectLabel.ForeColor = Color.FromArgb(187, 20, 20);
                NoConnectLabel.Text = "Фиксированный курс";
                _cbrDailyRates = true;
                _nbrbDailyRates = true;
            }
            else
            {
                if (!_rateExist)
                {
                    _eurrubCurrency = 0;
                    _usdrubCurrency = 0;
                    _eurusdCurrency = 0;
                    _eurbyrCurrency = 0;
                    _cbrDailyRates = OrdersManager.CBRDailyRates(CurrencyDateTimePicker.Value.Date, ref _eurrubCurrency, ref _usdrubCurrency);
                    _eurbyrCurrency = CurrencyConverter.NbrbDailyRates(DateTime.Now);
                    //NBRBDailyRates = OrdersManager.NBRBDailyRates(CurrencyDateTimePicker.Value.Date, ref EURBYRCurrency);
                    if (_usdrubCurrency != 0)
                        _eurusdCurrency = Decimal.Round(_eurrubCurrency / _usdrubCurrency, 4, MidpointRounding.AwayFromZero);
                }
            }

            if (currencyTypeId != 5)
                CurrencyTypeComboBox.SelectedIndex = currencyTypeId - 1;
            else
                CurrencyTypeComboBox.SelectedIndex = 3;

            if (discountPaymentConditionId == 1)
                rbDelayOfPayment.Checked = true;
            if (discountPaymentConditionId == 2)
                rbHalfOfPayment.Checked = true;
            if (discountPaymentConditionId == 3)
                rbFullPayment.Checked = true;
            if (discountPaymentConditionId == 5)
                kryptonRadioButton3.Checked = true;
            if (discountPaymentConditionId == 6)
                kryptonRadioButton5.Checked = true;
            if (discountPaymentConditionId == 4)
            {
                panel8.Visible = true;
                rbFactoring.Checked = true;
                string factoringName = _ordersManager.DiscountFactoringName(discountFactoringId);
                if (factoringName == kryptonRadioButton2.Text)
                    kryptonRadioButton2.Checked = true;
                if (factoringName == kryptonRadioButton4.Text)
                    kryptonRadioButton4.Checked = true;
                if (factoringName == kryptonRadioButton1.Text)
                    kryptonRadioButton1.Checked = true;
            }
            _rate = originalRate;
            CurrencyTextBox.Text = paymentRate.ToString();
            if (desireDate != DBNull.Value)
                DispatchDateTimePicker.Value = Convert.ToDateTime(desireDate);
            else
            {
                NoDateCheckBox.Checked = true;
            }
            CurrencyTypeComboBox_SelectionChangeCommitted(null, null);
        }

        private void Initialize()
        {
            //CbrBankData = OrdersManager.GetCbrBankData(CurrencyDateTimePicker.Value.Date);
            //NbrbBankData = OrdersManager.GetNbrbBankData(CurrencyDateTimePicker.Value.Date);

            //int DelayOfPayment = Convert.ToInt32(((DataRowView)OrdersManager.MegaOrdersBindingSource.Current).Row["DelayOfPayment"]);

            CurrencyTypeComboBox.SelectedIndex = 0;
            rbDelayOfPayment.Checked = true;
            tbDelayOfPayment.Visible = true;
            IsComplaintCheckBox.Checked = Convert.ToBoolean(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["IsComplaint"]);
            TransportCostTextEdit.Text = ((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["TransportCost"].ToString();
            AdditionalCostTextEdit.Text = ((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["AdditionalCost"].ToString();
            tbComplaintProfilCost.Text = ((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["ComplaintProfilCost"].ToString();
            tbComplaintTPSCost.Text = ((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["ComplaintTPSCost"].ToString();
            tbDelayOfPayment.Text = ((DataRowView)_ordersManager.MegaOrdersBindingSource.Current).Row["DelayOfPayment"].ToString();

            if (TransportCostTextEdit.Text.Length == 0)
                TransportCostTextEdit.Text = "0";
            if (AdditionalCostTextEdit.Text.Length == 0)
                AdditionalCostTextEdit.Text = "0";
            if (tbComplaintProfilCost.Text.Length == 0)
                tbComplaintProfilCost.Text = "0";
            if (tbComplaintTPSCost.Text.Length == 0)
                tbComplaintTPSCost.Text = "0";

            if (!_cbrDailyRates && !_nbrbDailyRates)
            {
                NoConnectLabel.ForeColor = Color.FromArgb(187, 20, 20);
                NoConnectLabel.Text = "Нет связи с банками";
            }
            else
                NoConnectLabel.Text = "Связь с банками установлена";

            if (_fixedPaymentRate)
            {
                _usdrubCurrency = 1;
                NoConnectLabel.ForeColor = Color.FromArgb(187, 20, 20);
                NoConnectLabel.Text = "Фиксированный курс";
                _cbrDailyRates = true;
                _nbrbDailyRates = true;
            }
            CurrencyTextBox.Text = "1";
        }

        private void CurrencyDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            //CbrBankData = "";
            //NbrbBankData = "";

            //CbrBankData = OrdersManager.GetCbrBankData(CurrencyDateTimePicker.Value.Date);
            //NbrbBankData = OrdersManager.GetNbrbBankData(CurrencyDateTimePicker.Value.Date);

            if (!_cbrDailyRates && !_nbrbDailyRates)
            {
                NoConnectLabel.ForeColor = Color.FromArgb(187, 20, 20);
                NoConnectLabel.Text = "Нет связи с банками";
            }
            else
                NoConnectLabel.Text = "Связь с банками установлена";

            if (_fixedPaymentRate)
            {
                _usdrubCurrency = 1;
                NoConnectLabel.ForeColor = Color.FromArgb(187, 20, 20);
                NoConnectLabel.Text = "Фиксированный курс";
                _cbrDailyRates = true;
                _nbrbDailyRates = true;
            }
            CurrencyTypeComboBox_SelectionChangeCommitted(null, null);
        }

        private void CurrencyTypeComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Евро")
                {
                    _rate = _eurbyrCurrency / BYN;
                    CurrencyTextBox.Text = _rate.ToString();
                    return;
                }
            }
            catch (DivideByZeroException exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show($"Параметр BYN равен {BYN}. Расчет не выполнен");
                return;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show("Ошибка. Расчет не выполнен");
                return;

            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Доллар США")
            {
                if (_eurrubCurrency == 0 || _usdrubCurrency == 0)
                {
                    NoConnectLabel.ForeColor = Color.FromArgb(187, 20, 20);
                    NoConnectLabel.Text = "Нет связи с банком";
                    return;
                }
                _rate = _eurusdCurrency;
                CurrencyTextBox.Text = _eurusdCurrency.ToString();

                if (CurrencyTextBox.Text != "0")
                {
                    NoConnectLabel.ForeColor = Color.Black;
                    NoConnectLabel.Text = "Данные получены с cbr.ru";
                }
                else
                {
                    NoConnectLabel.ForeColor = Color.FromArgb(187, 20, 20);
                    NoConnectLabel.Text = "Нет связи с банком";
                }
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Российский рубль")
            {
                _rate = _eurrubCurrency;
                CurrencyTextBox.Text = _eurrubCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurrubCurrency * 1.05m).ToString();

                if (CurrencyTextBox.Text != "0")
                {
                    NoConnectLabel.ForeColor = Color.Black;
                    NoConnectLabel.Text = "Данные получены с cbr.ru";
                }
                else
                {
                    NoConnectLabel.ForeColor = Color.FromArgb(187, 20, 20);
                    NoConnectLabel.Text = "Нет связи с банком";
                }
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Белорусский рубль")
            {
                _rate = _eurbyrCurrency;
                CurrencyTextBox.Text = _eurbyrCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurbyrCurrency * 1.05m).ToString();

                if (CurrencyTextBox.Text != "0")
                {
                    NoConnectLabel.ForeColor = Color.Black;
                    NoConnectLabel.Text = "Данные получены с nbrb.by";
                }
                else
                {
                    NoConnectLabel.ForeColor = Color.FromArgb(187, 20, 20);
                    NoConnectLabel.Text = "Нет связи с банком";
                }
            }
            if (_fixedPaymentRate)
            {
                _usdrubCurrency = 1;
                NoConnectLabel.ForeColor = Color.FromArgb(187, 20, 20);
                NoConnectLabel.Text = "Фиксированный курс";
                _cbrDailyRates = true;
                _nbrbDailyRates = true;
            }
        }

        private void NoDateCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (NoDateCheckBox.Checked == true)
                DispatchDateTimePicker.Enabled = false;
            else
                DispatchDateTimePicker.Enabled = true;
        }

        private void CurrencyForm_Load(object sender, EventArgs e)
        {

        }

        private void rbFactoring_CheckedChanged(object sender, EventArgs e)
        {
            panel8.Visible = true;
            tbDelayOfPayment.Visible = false;

            try
            {
                if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Евро")
                {
                    _rate = _eurbyrCurrency / BYN;
                    CurrencyTextBox.Text = _rate.ToString();
                    return;
                }
            }
            catch (DivideByZeroException exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show($"Параметр BYN равен {BYN}. Расчет не выполнен");
                return;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show("Ошибка. Расчет не выполнен");
                return;

            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Доллар США")
            {
                _rate = _eurusdCurrency;
                CurrencyTextBox.Text = _eurusdCurrency.ToString();
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Российский рубль")
            {
                _rate = _eurrubCurrency;
                CurrencyTextBox.Text = _eurrubCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurrubCurrency * 1.05m).ToString();
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Белорусский рубль")
            {
                _rate = _eurbyrCurrency;
                CurrencyTextBox.Text = _eurbyrCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurbyrCurrency * 1.05m).ToString();
            }
        }

        private void rbDelayOfPayment_CheckedChanged(object sender, EventArgs e)
        {
            panel8.Visible = false;
            tbDelayOfPayment.Visible = true;
            try
            {
                if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Евро")
                {
                    _rate = _eurbyrCurrency / BYN;
                    CurrencyTextBox.Text = _rate.ToString();
                    return;
                }
            }
            catch (DivideByZeroException exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show($"Параметр BYN равен {BYN}. Расчет не выполнен");
                return;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show("Ошибка. Расчет не выполнен");

            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Доллар США")
            {
                _rate = _eurusdCurrency;
                CurrencyTextBox.Text = _eurusdCurrency.ToString();
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Российский рубль")
            {
                _rate = _eurrubCurrency;
                CurrencyTextBox.Text = _eurrubCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurrubCurrency * 1.05m).ToString();
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Белорусский рубль")
            {
                _rate = _eurbyrCurrency;
                CurrencyTextBox.Text = _eurbyrCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurbyrCurrency * 1.05m).ToString();
            }
        }

        private void rbFullPayment_CheckedChanged(object sender, EventArgs e)
        {
            panel8.Visible = false;
            tbDelayOfPayment.Visible = false;
            try
            {
                if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Евро")
                {
                    _rate = _eurbyrCurrency / BYN;
                    CurrencyTextBox.Text = _rate.ToString();
                    return;
                }
            }
            catch (DivideByZeroException exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show($"Параметр BYN равен {BYN}. Расчет не выполнен");
                return;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show("Ошибка. Расчет не выполнен");
                return;

            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Доллар США")
            {
                _rate = _eurusdCurrency;
                CurrencyTextBox.Text = _eurusdCurrency.ToString();
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Российский рубль")
            {
                _rate = _eurrubCurrency;
                CurrencyTextBox.Text = _eurrubCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurrubCurrency * 1.05m).ToString();
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Белорусский рубль")
            {
                _rate = _eurbyrCurrency;
                CurrencyTextBox.Text = _eurbyrCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurbyrCurrency * 1.05m).ToString();
            }
        }

        private void rbHalfOfPayment_CheckedChanged(object sender, EventArgs e)
        {
            panel8.Visible = false;
            tbDelayOfPayment.Visible = false;
            try
            {
                if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Евро")
                {
                    _rate = _eurbyrCurrency / BYN;
                    CurrencyTextBox.Text = _rate.ToString();
                    return;
                }
            }
            catch (DivideByZeroException exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show($"Параметр BYN равен {BYN}. Расчет не выполнен");
                return;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show("Ошибка. Расчет не выполнен");
                return;

            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Доллар США")
            {
                _rate = _eurusdCurrency;
                CurrencyTextBox.Text = _eurusdCurrency.ToString();
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Российский рубль")
            {
                _rate = _eurrubCurrency;
                CurrencyTextBox.Text = _eurrubCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurrubCurrency * 1.05m).ToString();
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Белорусский рубль")
            {
                _rate = _eurbyrCurrency;
                CurrencyTextBox.Text = _eurbyrCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurbyrCurrency * 1.05m).ToString();
            }
        }

        private void IsComplaintCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            //if (IsComplaintCheckBox.Checked)
            //{
            //    tbComplaintProfilCost.Enabled = true;
            //    tbComplaintTPSCost.Enabled = true;
            //    tbComplaintNotes.Enabled = true;
            //}
            //else
            //{
            //    tbComplaintProfilCost.Enabled = false;
            //    tbComplaintTPSCost.Enabled = false;
            //    tbComplaintNotes.Enabled = false;
            //}
        }

        private void rbFCA_CheckedChanged(object sender, EventArgs e)
        {
            TransportCostTextEdit.Visible = false;
        }

        private void rbCpt_CheckedChanged(object sender, EventArgs e)
        {
            if (Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]) == 0
                || Convert.ToInt32(((DataRowView)_ordersManager.MegaOrdersBindingSource.Current)["AgreementStatusID"]) == 3)
            {
                _transportCost = _ordersManager.TransportCost(_weight);
            }
            TransportCostTextEdit.Text = _transportCost.ToString();
            TransportCostTextEdit.Visible = true;
        }

        private void TransportCostTextEdit_TextChanged(object sender, EventArgs e)
        {

        }

        private void kryptonRadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            panel8.Visible = false;
            tbDelayOfPayment.Visible = false;

            try
            {
                if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Евро")
                {
                    _rate = _eurbyrCurrency / BYN;
                    CurrencyTextBox.Text = _rate.ToString();
                    return;
                }
            }
            catch (DivideByZeroException exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show($"Параметр BYN равен {BYN}. Расчет не выполнен");
                return;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show("Ошибка. Расчет не выполнен");
                return;

            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Доллар США")
            {
                _rate = _eurusdCurrency;
                CurrencyTextBox.Text = _eurusdCurrency.ToString();
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Российский рубль")
            {
                _rate = _eurrubCurrency;
                CurrencyTextBox.Text = _eurrubCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurrubCurrency * 1.05m).ToString();
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Белорусский рубль")
            {
                _rate = _eurbyrCurrency;
                CurrencyTextBox.Text = _eurbyrCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurbyrCurrency * 1.05m).ToString();
            }
        }

        private void kryptonRadioButton5_CheckedChanged(object sender, EventArgs e)
        {
            panel8.Visible = false;
            tbDelayOfPayment.Visible = false;

            try
            {
                if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Евро")
                {
                    _rate = _eurbyrCurrency / BYN;
                    CurrencyTextBox.Text = _rate.ToString();
                    return;
                }
            }
            catch (DivideByZeroException exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show($"Параметр BYN равен {BYN}. Расчет не выполнен");
                return;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show("Ошибка. Расчет не выполнен");
                return;

            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Доллар США")
            {
                _rate = _eurusdCurrency;
                CurrencyTextBox.Text = _eurusdCurrency.ToString();
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Российский рубль")
            {
                _rate = _eurrubCurrency;
                CurrencyTextBox.Text = _eurrubCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurrubCurrency * 1.05m).ToString();
            }

            if (CurrencyTypeComboBox.SelectedItem.ToString() == "Евро - Белорусский рубль")
            {
                _rate = _eurbyrCurrency;
                CurrencyTextBox.Text = _eurbyrCurrency.ToString();
                if (_clientId != 145 && rbDelayOfPayment.Checked)
                    CurrencyTextBox.Text = (_eurbyrCurrency * 1.05m).ToString();
            }
        }
    }
}
