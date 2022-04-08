using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Configuration;
// System.Data необходимо для использования DataSet
using System.Data;
// System.Data.OleDb необходимо для использования OleDbConnection.
using System.Data.OleDb;
// using System.IO; необходимо для использования OpenFileDialog
using System.ComponentModel;
using System.IO;
using Microsoft.Win32;

namespace MSAccessClient
{
    /// <summary>
    /// lamer spagghetti logic for MainWindow
    /// </summary>
    public partial class MainWindow : Window
    {
        private string dbPath = null;
        private string conString = null;
        private string tableName = "ШОП";
        private string pkColumnName = "ID";
        private bool bAdapterConfigured = false;
        private int lastChangedRowIndex = -1;
        private List<int> inttype_col_ind = new List<int>() { 2, 4, 8, 9 };
        private List<int> notnull_col_ind = new List<int>() { };
        private List<int> defaultnull_col_ind = new List<int>() { 2, 4, 8, 9, 58};
        private List<int> cmb_col_ind = new List<int>() { 3, 4, 5, 12, 58 };     // позиции колонок с ComboBox в DataGrid 
        private List<int> chb_col_ind = new List<int>() { 13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,  // позиции колонок с CheckBox DataGrid
            28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57};       
        private int spBlockHeight = 25;
        private int spBlockWidth = 150;
        private const int p = 100;
        private int oledbdrparamsdatecounter = 1;
        private List<string> comparison_items1 = new List<string>() { "Больше", "Ровно", "Меньше" };
        private List<string> comparison_items2 = new List<string>() { "Позже", "В этот день", "Раньше" };
        private List<string> comparison_signs1 = new List<string>() { ">", "=", "<" };
        private List<string> gender_items = new List<string>() { "М", "Ж" /* другие - не люди */ };
        private int cur_dg_cmb_id = -1;
        private string sp1_age_cmb1_text = "";
        private string sp1_age_tbx1_text = "";
        private string sp1_fio_tbx1_text = "";
        private string sp1_injurydate_cmb1_text = "";
        private string sp1_injurydate_tbx1_text = "";
        private string sp1_datemed_cmb1_text = "";
        private string sp1_datemed_tbx1_text = "";
        private string sp1_daysbeforemed_cmb1_text = "";
        private string sp1_daysbeforemed_tbx1_text = "";
        private string sp1_daystreat000_cmb1_text = "";
        private string sp1_daystreat001_cmb1_text = "";
        private string sp1_daystreat002_cmb1_text = "";
        private string sp1_daystreat000_tbx1_text = "";
        private string sp1_daystreat001_tbx1_text = "";
        private string sp1_daystreat002_tbx1_text = "";
        private List<string> regexpert_items = new List<string>() { "1. Брест", "2. Витебск", "3. Гомель", "4. Гродно", "5. Минская обл.", "6. Могилев", "7. Минск" };
        private List<string> regtypeexpert_items = new List<string>() { "Город", "Деревня, село, поселок", "Поселок гор. типа", "Агрогородок", "Другое" };
        private List<string> injurytype_items = new List<string>() { "Транспортная", "Падение на пл-ти", "Падение с высоты", "Ныряние", "Удар в голову", "Прямая уд. нагр.", "Другое" };
        private List<string> pain_items = new List<string>() { "Шея", "Голова", "Туловище", "Конечности" };
        private List<string> dizziness_items = new List<string>() { "Есть головокружение", "Нет головокружения" };
        private List<string> nausea_items = new List<string>() { "Есть тошнота", "Нет тошноты" };
        private List<string> fulltiredness_items = new List<string>() { "Есть общ. слабость", "Нет общ. слабости" };
        private List<string> tiredness_items = new List<string>() { "Верх. конечностей", "Нижн. конечностей" };
        private List<string> numbness_items = new List<string>() { "Верх. конечностей", "Нижн. конечностей" };
        private List<string> goosebumps_items = new List<string>() { "Верх. конечностей", "Нижн. конечностей", "Туловища" };
        private List<string> nfto1_items = new List<string>() { "Есть НФТО (суб)", "Нет НФТО (суб)" };
        private List<string> bodydmg_items = new List<string>() { "Голова", "Шея", "Туловище", "Верх. конечности", "Нижн. конечности" };
        private List<string> sensivitydmg_items = new List<string>() { "Голова", "Шея", "Туловище", "Верх. конечности", "Нижн. конечности" };
        private List<string> limbsweakness_items = new List<string>() { "Верх. конечности", "Нижн. конечности" };
        private List<string> nfto2_items = new List<string>() { "Есть НФТО (об)", "Нет НФТО (об)" };
        private List<string> neckdmg_items = new List<string>() { "Есть мыш. деф. шеи", "Нет мыш. деф. шеи" };
        private List<string> mrt_items = new List<string>() { "МРТ вып.", "МРТ подтв." };
        private List<string> mskt_items = new List<string>() { "МСКТ вып.", "МСКТ подтв." };
        private List<string> xray_items = new List<string>() { "Рентген вып.", "Рентген подтв." };
        private List<string> emg_items = new List<string>() { "ЭМГ вып.", "ЭМГ подтв." };
        private List<string> uzi_items = new List<string>() { "УЗИ вып.", "УЗИ подтв." };
        private List<string> eeg_items = new List<string>() { "ЭЭГ вып.", "ЭЭГ подтв." };
        private List<string> rvg_items = new List<string>() { "РВГ вып.", "РВГ подтв." };
        private List<string> drugoi_items = new List<string>() { "Другой вып.", "Другой подтв." };
        private List<string> alcohol_items = new List<string>() { "0", "0.01 - 0.29 ‰", "0.3 - 1.49 ‰", "1.5 - 2.49 ‰", "2.5 - 2.99 ‰", "3.0 - 4.99 ‰", "5.0 ‰ и больше", "Нет данных" };
        private List<string> injurytype_columns = new List<string>() { "ВидТравмыТранспортная", "ВидТравмыПадениеПлоск", "ВидТравмыПадениеВыс", "ВидТравмыНыряние", "ВидТравмыУдарВГолову", "ВидТравмыУдарнаяНагр", "ВидТравмыДругое"};
        private List<string> pain_columns = new List<string>() { "БольШея", "БольГолова", "БольТуловище", "БольКонечности" };
        private List<string> tiredness_columns = new List<string>() { "СлабостьВерхКон", "СлабостьНижКон" };
        private List<string> numbness_columns = new List<string>() { "ОнемениеВерхКон", "ОнемениеНижКон" };
        private List<string> goosebumps_columns = new List<string>() { "ЧувМурашекВерхКон", "ЧувМурашекНижКон", "ЧувМурашекТуловище" };
        private List<string> bodydmg_columns = new List<string>() { "НаружТелПоврГолова", "НаружТелПоврШея", "НаружТелПоврТуловище", "НаружТелПоврВерхКон", "НаружТелПоврНижКон" };
        private List<string> sensivitydmg_columns = new List<string>() { "НарушЧувГолова", "НарушЧувШея", "НарушЧувТуловище", "НарушЧувВерхКон", "НарушЧувНижКон" };
        private List<string> limbsweakness_columns = new List<string>() { "УменьшСилыВерхКон", "УменьшСилыНижКон"};
        private List<string> mrt_columns = new List<string>() { "МРТвып", "МРТподтв" };
        private List<string> mskt_columns = new List<string>() { "МСКТвып", "МСКТподтв" };
        private List<string> xray_columns = new List<string>() { "РЕНТГЕНВып", "РЕНТГЕНподтв" };
        private List<string> emg_columns = new List<string>() { "ЭМГвып", "ЭМГподтв" };
        private List<string> uzi_columns = new List<string>() { "УЗИвып", "УЗИподтв" };
        private List<string> eeg_columns = new List<string>() { "ЭЭГвып", "ЭЭГподтв" };
        private List<string> rvg_columns = new List<string>() { "РВГвып", "РВГподтв" };
        private List<string> drugoi_columns = new List<string>() { "ДРУГОЕвып", "ДРУГОЕподтв" };
        private List<bool> sp1_gender_chb_ischecked = new List<bool>( new bool[2]);
        private List<bool> sp1_regexpert_chb_ischecked = new List<bool>(new bool[7]);
        private List<bool> sp1_regtypeexpert_chb_ischecked = new List<bool>( new bool[5]);
        private List<bool> sp1_injurytype_chb_ischecked = new List<bool>(new bool[7]);
        private List<bool> sp1_pain_chb_ischecked = new List<bool>(new bool[4]);
        private List<bool> sp1_dizziness_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_nausea_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_fulltiredness_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_tiredness_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_numbness_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_goosebumps_chb_ischecked = new List<bool>(new bool[3]);
        private List<bool> sp1_nfto1_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_bodydmg_chb_ischecked = new List<bool>(new bool[5]);
        private List<bool> sp1_sensivitydmg_chb_ischecked = new List<bool>(new bool[5]);
        private List<bool> sp1_limbsweakness_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_nfto2_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_neckdmg_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_mrt_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_mskt_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_xray_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_emg_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_uzi_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_eeg_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_rvg_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_drugoi_chb_ischecked = new List<bool>(new bool[2]);
        private List<bool> sp1_alcohol_chb_ischecked = new List<bool>(new bool[8]);
        private string[] sp1_cat_operators = new string[p];
        private DBCon dbCon = null;
        private DataSet ds;
        private OleDbDataAdapter dsAdapter;

        public MainWindow()
        {
            InitializeComponent();

            cmb_gender.ItemsSource = gender_items;
            cmb_regexpert.ItemsSource = regexpert_items;
            cmb_regtypeexpert.ItemsSource = regtypeexpert_items;
            cmb_injurytype.ItemsSource = injurytype_items;
            cmb_alcohol.ItemsSource = alcohol_items;
        }

        // вспомогательные методы:
        private void ConfigDS()
        {
            try
            {
                ds = new DataSet();

                // Заполнение DataSet'а и DataGrid'а:
                dsAdapter.Fill(ds);
                dg.ItemsSource = ds.Tables[0].DefaultView;
                dg.AutoGenerateColumns = false;


                // Создание ограничения числового первичного ключа с AUTO_INCREMENT:
                int lastID = 0;
                if(ds.Tables[0].Rows.Count > 0)
                {
                    ds.Tables[0].PrimaryKey = new DataColumn[] { ds.Tables[0].Columns[pkColumnName] };
                    ds.Tables[0].Columns[pkColumnName].AutoIncrement = true;
                    var temp = (from m in ds.Tables[0].AsEnumerable() select m[pkColumnName]).Max();
                    lastID = (int)temp;
                }            
                
                ds.Tables[0].Columns[pkColumnName].AutoIncrementSeed = lastID + 1;
                ds.Tables[0].Columns[pkColumnName].AutoIncrementStep = 1;

                foreach (int i in chb_col_ind) { ds.Tables[0].Columns[i].DefaultValue = false; }
                foreach (int i in notnull_col_ind) { ds.Tables[0].Columns[i].AllowDBNull = false; }
                foreach (int i in defaultnull_col_ind) { ds.Tables[0].Columns[i].DefaultValue = DBNull.Value; }

                /*
                 * Замечание: при внесении в БД новой i-ой добавленной в клиенте строки в БД OleDbDataAdapter изменяет  
                 * ее ID на max + i, где max - это максимальный ID среди когда-либо добавленных строк в эту таблицу,
                 * даже если такая строка уже не существует в БД. 
                 * (особенность Access'а)   
                 */

                // Сортировка по первичному ключу :
                ListSortDirection sortDirection = ListSortDirection.Ascending;
                var column = dg.Columns[0];

                dg.Items.SortDescriptions.Clear();
                dg.Items.SortDescriptions.Add(new SortDescription(column.SortMemberPath, sortDirection));

                foreach (var col in dg.Columns)
                {
                    col.SortDirection = null;
                }
                column.SortDirection = sortDirection;
                dg.Items.Refresh();
            }
            catch (OleDbException ex)
            {
                dbPath = null; conString = null; dbCon = null;
                MessageBox.Show("Исключение типа OleDbException в ButtonClick_SelectDB(): " +
                    "Возможно, другой пользователь (или вы) находится в конструкторе таблицы. Закройте Access и попробуйте снова." + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в ConfigDS(): " + ex.GetType().ToString() + " " + ex.Message);
            }
        }
        private void ConfigDSAdapter()
        {
            try
            {
                bAdapterConfigured = false;
                ds = new DataSet();
                dsAdapter = new OleDbDataAdapter("SELECT * FROM " + tableName, dbCon.getConnection());

                // Настройка комманд адаптера:
                OleDbCommandBuilder cb = new OleDbCommandBuilder(dsAdapter);
                dsAdapter.InsertCommand = cb.GetInsertCommand();
                dsAdapter.UpdateCommand = cb.GetUpdateCommand();
                dsAdapter.DeleteCommand = cb.GetDeleteCommand();

                bAdapterConfigured = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в ConfigDSAdapter(): " + ex.GetType().ToString() + " " + ex.Message);
            }
        }
        private void Window_FillStackPanelCategorical(ref StackPanel sp, ref StackPanel sp_chb, ref CheckBox cb, ref List<bool> chb_ischecked, Button b, List<string> items, int blockHeight, int blockWidth, int methodid, bool isBinary, bool multichoice)
        {
            try
            {
                sp.Children.Clear();
                sp_chb.Children.Clear();
                int k = items.Count;

                if (cb.IsChecked == true)
                {
                    cb.IsChecked = false;
                    List<TextBlock> tbk = new List<TextBlock>(k); for (int i = 0; i < k; i++) { tbk.Add(new TextBlock() { Height = blockHeight, Width = blockWidth, Text = items[i] }); };

                    for (int i = 0; i < k; i++) { sp.Children.Add(tbk[i]); };
                    sp.Children.Add(b);

                    List<CheckBox> chb = new List<CheckBox>(k); for (int i = 0; i < k; i++) { chb.Add(new CheckBox() { Height = blockHeight, Width = blockHeight /* не опечатка: checkbox квадратный */, IsChecked = false }); };
                    for (int i = 0; i < k; i++) { sp_chb.Children.Add(chb[i]); };

                    if (isBinary)
                    {
                        sp_chb.Children.Add(new TextBlock() { Height = blockHeight, Width = 80, Text = "И (&&)" });
                    }
                    else
                    {
                        if(multichoice)
                        {
                            ComboBox and_or = new ComboBox() { Height = blockHeight, Width = 80, HorizontalAlignment = HorizontalAlignment.Left };
                            and_or.Items.Add("И (&&)"); and_or.Items.Add("ИЛИ (||)");
                            sp_chb.Children.Add(and_or);
                        }
                        else
                        {
                            sp_chb.Children.Add(new TextBlock() { Height = blockHeight, Width = 80, Text = "ИЛИ (||)" });
                        }                              
                    }      
                }
                else
                {
                    chb_ischecked = new List<bool>(new bool[k]);
                    cb.IsChecked = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в Window_OnSP1CheckBox" + methodid.ToString() + "Clicked() -> Window_FillStackPanelCategorical(): " + ex.GetType().ToString() + " " + ex.Message);
            }
        }
        private void Window_FillStackPanelScalar(ref StackPanel sp, ref StackPanel sp_chb, ref CheckBox cb, ref string cmbtext_field_toreset, ref string tbxtext_field_toreset, Button b, List<string> comparison_items, string tbx_tag, int blockHeight, int blockWidth, int methodid)
        {
            try
            {                
                sp.Children.Clear();
                sp_chb.Children.Clear();

                int k = comparison_items.Count;


                if (cb.IsChecked == true)
                {
                    cb.IsChecked = false;
                    ComboBox cmb1 = new ComboBox() { Height = blockHeight, Width = blockWidth, HorizontalAlignment = HorizontalAlignment.Left };
                    for (int i = 0; i < k; i++) { cmb1.Items.Add(comparison_items[i]); }; 

                    TextBox tbx1 = new TextBox() { Height = blockHeight, Width = blockWidth, Style = this.Resources["MyWaterMarkStyle"] as Style, Tag = tbx_tag, HorizontalAlignment = HorizontalAlignment.Left };                                  

                    sp1.Children.Add(cmb1);
                    sp1.Children.Add(tbx1);
                    sp1.Children.Add(b);
                }
                else
                {
                    cb.IsChecked = false;
                    cmbtext_field_toreset = "";
                    tbxtext_field_toreset = "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в Window_OnSP1CheckBox" + methodid.ToString() + "Clicked() -> Window_FillStackPanelScalar(): " + ex.GetType().ToString() + " " + ex.Message);
            }
        }
        private void Window_FillStackPanelSingleInput(ref StackPanel sp, ref StackPanel sp_chb, ref CheckBox cb, ref string tbxtext_field_toreset, Button b, string tbx_tag, int blockHeight, int blockWidth, int methodid)
        {
            try
            {
                sp1.Children.Clear();
                sp1_chb.Children.Clear();
                
                if (cb.IsChecked == true)
                {
                    cb.IsChecked = false;
                    TextBox tbx1 = new TextBox() { Height = blockHeight, Width = blockWidth, Style = this.Resources["MyWaterMarkStyle"] as Style, Tag = tbx_tag, HorizontalAlignment = HorizontalAlignment.Left };

                    sp.Children.Add(tbx1);
                    sp.Children.Add(b);
                }
                else
                {
                    tbxtext_field_toreset = "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в Window_OnSP1CheckBox" + methodid.ToString() + "Clicked() -> Window_FillStackPanelSingleInput(): " + ex.GetType().ToString() + " " + ex.Message);
            }
        }
        private void Window_ConfirmStackPanelCategorical(ref StackPanel sp, ref StackPanel sp_chb, ref CheckBox cb, ref List<bool> chb_ischecked, int methodid, bool isBinary, bool multichoice)
        {
            try
            {
                int k = sp_chb.Children.Count - 1;
                int flag = 0;

                if (isBinary)
                {
                    flag = 1; sp1_cat_operators[methodid] = "AND";
                }
                else
                {
                    if(multichoice)
                    {
                        ComboBox cmb = sp_chb.Children[k] as ComboBox;
                        if (cmb.SelectedIndex > -1)
                        {
                            flag = 1; if (cmb.SelectedIndex == 0) { sp1_cat_operators[methodid] = "AND"; } else { sp1_cat_operators[methodid] = "OR"; }
                        }
                        else
                        {
                            MessageBox.Show("Логический опратор (И / ИЛИ) не выбран, попробуйте снова.");
                        }
                    }
                    else
                    {
                        flag = 1; sp1_cat_operators[methodid] = "OR";
                    }   
                }
                
                if (flag == 1)
                {
                    List<CheckBox> chb = new List<CheckBox>(); for (int i = 0; i < k; i++) { chb.Add(sp_chb.Children[i] as CheckBox); };
                    chb_ischecked = new List<bool>(new bool[k]); // обнуление (превр. в список из false) 

                    for (int i = 0; i < k; i++) { chb_ischecked[i] = chb[i].IsChecked.HasValue ? chb[i].IsChecked.Value : false; };

                    if (chb_ischecked.Contains(true) && chb_ischecked.Contains(false))
                    {                        
                        cb.IsChecked = true;
                    }
                    else
                    {
                        if (isBinary)
                        {
                            cb.IsChecked = false;
                            chb_ischecked = new List<bool>(new bool[k]);  // обнуление (превр. в список из false)                            
                        }                      
                        else
                        {
                            //if(temp == "OR") { cb.IsChecked = false; chb_ischecked = new List<bool>(new bool[k]); } else { cb.IsChecked = true; }                            
                            cb.IsChecked = chb_ischecked.Contains(true);
                        }
                    }
                } 
               
                sp.Children.Clear();
                sp_chb.Children.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в ButtonClick_SP1Ok" + methodid.ToString() + "() -> Window_ConfirmStackPanelCategorical(): " + ex.GetType().ToString() + " " + ex.Message);
            }
        }
        private void Window_ConfirmStackPanelScalar(ref StackPanel sp, ref StackPanel sp_chb, ref CheckBox cb, ref string cmbtext_field_toset, ref string tbxtext_field_toset, List<string> comparison_items, List<string> signs, bool isDateTimetype, bool isInteger, int methodid)
        {
            try
            {
                ComboBox cmb1 = sp.Children[0] as ComboBox;
                TextBox tbx1 = sp.Children[1] as TextBox;
                int k = comparison_items.Count;
                bool correctdata = false;

                if (cmb1.Text == "")
                {
                    cb.IsChecked = false;
                    MessageBox.Show("Вы не выбрали оператор сравнения, попробуйте снова. ");
                }
                else
                {
                    if(isDateTimetype)
                    {
                        DateTime temp1;
                        correctdata = DateTime.TryParse(tbx1.Text, out temp1);
                    }
                    else
                    {
                        if(isInteger)
                        {
                            int temp2;
                            correctdata = int.TryParse(tbx1.Text, out temp2);
                        }
                        else
                        {
                            double temp3;
                            correctdata = double.TryParse(tbx1.Text, out temp3);
                        }
                    }
                    
                    if (correctdata)
                    {
                        for(int i = 0; i < k; i++)
                        {                           
                            if (cmb1.Text == comparison_items[i])
                            {                               
                                cmbtext_field_toset = signs[i];
                                tbxtext_field_toset = tbx1.Text;
                                cb.IsChecked = true;
                                break;
                            }
                        }
                    }
                    else
                    {
                        if(isDateTimetype)
                        {
                            MessageBox.Show("Введенная строка не представляет правильную дату. Попробуйте формат 'дд.мм.гггг' без пробелов.");
                        }
                        else
                        {
                            if (isInteger)
                            {
                                MessageBox.Show("Введенная строка не представляет собой целое число.");
                            }
                            else
                            {
                                MessageBox.Show("Введенную строку не получилось преобразовать в действительное число. Попробуйте отделить дробную часть от целой с помощью ',' (запятой).");
                            }
                        }                                              
                    }
                }

                sp.Children.Clear();
                sp_chb.Children.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в ButtonClick_SP1Ok" + methodid.ToString() + "() -> Window_ConfirmStackPanelScalar(): " + ex.GetType().ToString() + " " + ex.Message);
            }
        }
        private void Window_ConfirmStackPanelSingleInput(ref StackPanel sp, ref StackPanel sp_chb, ref CheckBox cb, ref string tbxtext_field_toset, int methodid)
        {
            try
            {
                TextBox tbx1 = sp.Children[0] as TextBox;
                tbxtext_field_toset = tbx1.Text;

                if (tbxtext_field_toset == "")
                {
                    cb.IsChecked = false;
                    MessageBox.Show("Введена пустая строка. Требование проигнорировано.");
                }
                else
                {
                    cb.IsChecked = true;
                }
                sp.Children.Clear();
                sp_chb.Children.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в ButtonClick_SP1Ok" + methodid.ToString() + "() -> Window_ConfirmStackPanelSingleInput(): " + ex.GetType().ToString() + " " + ex.Message);
            }
        }
        private string BuildCommandStringHelperSingleCase(ref OleDbCommand oledbcom, ref CheckBox chb, ref bool flag, string colname, int mycase, string inptext, string inpop)
        {
            string str = "";
            string res = "";

            try
            {
                if (chb.IsChecked.HasValue ? chb.IsChecked.Value : false)
                {
                    if (mycase == 1) { str = colname + " " + inpop + " " + "'%" + inptext + "%'"; }  // строка
                    if (mycase == 2) { str = colname + " " + inpop + " " + inptext; }                // число
                    if (mycase == 3)
                    {
                        string paramname = "@date" + oledbdrparamsdatecounter.ToString();
                        oledbcom.Parameters.Add(paramname, OleDbType.Date).Value = DateTime.Parse(inptext);
                        str = colname + " " + inpop + " " + paramname;
                        oledbdrparamsdatecounter++;
                    }
                    res = " ( " + str + " ) AND";
                    flag = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в BuildCommandStringHelperSingleCase(): " + ex.GetType().ToString() + "  " + ex.Message);
            }

            return res;
        }
        private string BuildCommandStringHelperMultiCatCase(ref CheckBox chb, ref bool flag, string colname, List<bool> chb_values, List<string> chb_items, string op, bool binary, bool inttype)
        {
            string str = "";
            string res = "";

            try
            {
                if (binary)  // строка
                {
                    op = "AND";
                    if (chb.IsChecked.HasValue ? chb.IsChecked.Value : false)
                    {
                        for (int i = 0; i < chb_items.Count; i++)
                        {
                            if (chb_values[i])
                            {
                                str = colname + " = ";
                                if (inttype)
                                {
                                    str = str + int.Parse(chb_items[i]).ToString();
                                }
                                else
                                {
                                    str = str + "'" + chb_items[i] + "' ";
                                }
                            }
                        }
                        if (str != "")
                        {
                            res = " ( " + str + ") AND";
                            flag = true;
                        }
                    }
                }
                else
                {
                    if (chb.IsChecked.HasValue ? chb.IsChecked.Value : false)
                    {
                        for (int i = 0; i < chb_items.Count; i++)
                        {
                            if (chb_values[i])
                            {
                                if (inttype)
                                {
                                    str = str + colname + " = " + int.Parse(chb_items[i].Substring(0, 1)) + " " + op + " ";
                                }
                                else
                                {
                                    str = str + colname + " = " + "'" + chb_items[i] + "'" + " " + op + " ";
                                }
                            }
                        }
                        if (str != "")
                        {
                            int k = str.Length;
                            str = str.Substring(0, k - (op.Length + 2));
                            res = " ( " + str + " ) AND";
                            flag = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в BuildCommandStringHelperMultiCatCase() -> " + colname + ": " + ex.GetType().ToString() + "  " + ex.Message);
            }

            return res;
        }
        private string BuildCommandStringHelperSingleBoolCase(ref CheckBox chb, ref bool flag, string colname, List<bool> chb_values)
        {
            string str = "";
            string res = "";

            try
            {
                if (chb.IsChecked.HasValue ? chb.IsChecked.Value : false)
                {
                    if (chb_values[0]) { str = colname + " = " + "True "; }
                    else { str = colname + " = " + "False "; }

                    if (str != "")
                    {
                        int k = str.Length;

                        res = " ( " + str + " ) AND";
                        flag = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в BuildCommandStringHelperSingleBoolCase() " + ": " + ex.GetType().ToString() + "  " + ex.Message);
            }

            return res;
        }
        private string BuildCommandStringHelperMultiBoolCase(ref CheckBox chb, ref bool flag, List<string> colnames, List<bool> chb_values, string op)
        {
            string str = "";
            string res = "";

            try
            {
                if (chb.IsChecked.HasValue ? chb.IsChecked.Value : false)
                {
                    for (int i = 0; i < colnames.Count; i++)
                    {
                        if (chb_values[i])
                        {
                            str = str + colnames[i] + " = " + "True" + " " + op + " ";
                        }
                    }
                    if (str != "")
                    {
                        int k = str.Length;
                        str = str.Substring(0, k - (op.Length + 2));
                        res = " ( " + str + " ) AND";
                        flag = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в BuildCommandStringHelperMultiBoolCase() " + ": " + ex.GetType().ToString() + "  " + ex.Message);
            }

            return res;
        }
        
        // обработчики событий:
        private void Window_OnDataGridCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            try
            {
                if (e.EditAction == DataGridEditAction.Commit)
                {
                    lastChangedRowIndex = e.Row.GetIndex();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в Window_OnDataGridCellEditEnding(): " + ex.Message);
            }          
        }
        private void Window_OnDataGrid2AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.PropertyType == typeof(System.DateTime))
            {
                (e.Column as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";
            }        
            /*
            //Если колонка представляет собой первичный ключ, то скроем
            if (e.Column.Header.ToString() == pkColumnName)
            {
                e.Column.Visibility = Visibility.Hidden;
            }
            */
        }
        private void Window_OnDataGridComboBoxDropDownClosed(object sender, EventArgs e)
        {
            try
            {
                ComboBox cmb = (ComboBox)sender;
                int l = cmb.Name.Length;
                int dg_cmb_id = int.Parse(cmb.Name.Substring(l - 3, 3));
                int ind = -1;
                int flag = 0;   // флаг указывает, найден ли уже нужный обработчик или еще нет
                bool notempty = cmb.Text != "";
                cur_dg_cmb_id = dg_cmb_id;
                
                if (notempty)
                {
                    DataRowView drv = (DataRowView)dg.SelectedItems[0];
                    ind = ds.Tables[0].Rows.IndexOf(drv.Row);

                    if (flag == 0 && dg_cmb_id == 1 && cmb.Text != "Нет данных")  // РегионВыпЭксп
                    {
                        flag = 1;
                        if (cmb.Text != "Нет данных")
                        {
                            int newval = int.Parse(cmb.Text.Substring(0, 1));
                            ds.Tables[0].Rows[ind][cmb_col_ind[1]] = newval;
                        }
                        else
                        {
                            ds.Tables[0].Rows[ind][cmb_col_ind[1]] = DBNull.Value;
                        }
                    }

                    if (flag == 0 && cmb.Text != "")  // Остальные ComboBox'ы
                    {
                        flag = 1;
                        if(cmb.Text != "Нет данных")
                        {
                            ds.Tables[0].Rows[ind][cmb_col_ind[dg_cmb_id]] = cmb.Text;

                        }
                        else
                        {
                            ds.Tables[0].Rows[ind][cmb_col_ind[dg_cmb_id]] = DBNull.Value;
                        }
                    }
                }      
            }
            catch (IndexOutOfRangeException ex)
            {
                // ничего не делать, а остальные исключения ловить
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в Window_OnDataGridComboBoxDropDownClosed(): " + cur_dg_cmb_id.ToString()+ ": " + ex.GetType().ToString() + " " + ex.Message);
            }
        }      
        private void Window_OnDataGridCheckBoxClicked(object sender, EventArgs e)
        {
            try
            {          
                CheckBox chb = (CheckBox)sender;
                int l = chb.Name.Length;
                int dg_chb_id = int.Parse(chb.Name.Substring(l - 3, 3));

                DataRowView drv = (DataRowView)dg.SelectedItems[0];
                int ind = ds.Tables[0].Rows.IndexOf(drv.Row);
                if (ind != -1)
                {
                    ds.Tables[0].Rows[ind][chb_col_ind[dg_chb_id]] = chb.IsChecked;
                }                                               
            }
            catch (InvalidCastException ex)
            {
                // ничего не делать, а остальные исключения ловить 
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в Window_OnDataGridCheckBoxClicked(): " + ex.GetType().ToString() + " " + ex.Message);
            }            
        }
        private void Window_OnSP1CheckBox1Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = 70, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok1;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_gender, ref sp1_gender_chb_ischecked, b, gender_items, spBlockHeight, spBlockWidth, 1, true, false);
        }
        private void Window_OnSP1CheckBox2Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = 70, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok2;
            Window_FillStackPanelSingleInput(ref sp1, ref sp1_chb, ref sp1_chb_fio, ref sp1_fio_tbx1_text, b, "Искать эту подстроку...", spBlockHeight, spBlockWidth, 2);
        }
        private void Window_OnSP1CheckBox3Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = 70, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok3;
            Window_FillStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_age, ref sp1_age_cmb1_text, ref sp1_age_tbx1_text, b, comparison_items1, "Возраст (лет) ...", spBlockHeight, spBlockWidth, 3);
        }
        private void Window_OnSP1CheckBox4Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = 70, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok4;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_regexpert, ref sp1_regexpert_chb_ischecked, b, regexpert_items, spBlockHeight, spBlockWidth, 4, false, false);
        }
        private void Window_OnSP1CheckBox5Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok5;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_regtypeexpert, ref sp1_regtypeexpert_chb_ischecked, b, regtypeexpert_items, spBlockHeight, spBlockWidth, 5, false, false);
        }
        private void Window_OnSP1CheckBox6Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok6;
            Window_FillStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_injurydate, ref sp1_injurydate_cmb1_text, ref sp1_injurydate_tbx1_text, b, comparison_items2, "дд.мм.гггг", spBlockHeight, spBlockWidth,6);
        }
        private void Window_OnSP1CheckBox7Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok7;
            Window_FillStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_datemed, ref sp1_datemed_cmb1_text, ref sp1_datemed_tbx1_text, b, comparison_items2, "дд.мм.гггг", spBlockHeight, spBlockWidth, 7);
        }
        private void Window_OnSP1CheckBox8Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok8;
            Window_FillStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_daysbeforemed, ref sp1_daysbeforemed_cmb1_text, ref sp1_daysbeforemed_tbx1_text, b, comparison_items1, "дней", spBlockHeight, spBlockWidth, 8);
        }
        private void Window_OnSP1CheckBox9Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok9;
            Window_FillStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_daystreat000, ref sp1_daystreat000_cmb1_text, ref sp1_daystreat000_tbx1_text, b, comparison_items1, "дней", spBlockHeight, spBlockWidth, 9);
        }
        private void Window_OnSP1CheckBox10Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok10;
            Window_FillStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_daystreat001, ref sp1_daystreat001_cmb1_text, ref sp1_daystreat001_tbx1_text, b, comparison_items1, "дней", spBlockHeight, spBlockWidth, 10);
        }
        private void Window_OnSP1CheckBox11Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok11;
            Window_FillStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_daystreat002, ref sp1_daystreat002_cmb1_text, ref sp1_daystreat002_tbx1_text, b, comparison_items1, "дней", spBlockHeight, spBlockWidth, 11);
        }
        private void Window_OnSP1CheckBox12Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok12;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_injurytype, ref sp1_injurytype_chb_ischecked, b, injurytype_items, spBlockHeight, spBlockWidth, 12, false, false);
        }
        private void Window_OnSP1CheckBox13Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok13;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_pain, ref sp1_pain_chb_ischecked, b, pain_items, spBlockHeight, spBlockWidth, 13, false, true);
        }
        private void Window_OnSP1CheckBox14Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok14;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_dizziness, ref sp1_dizziness_chb_ischecked, b, dizziness_items, spBlockHeight, spBlockWidth, 14, true, false);
        }
        private void Window_OnSP1CheckBox15Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok15;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_nausea, ref sp1_nausea_chb_ischecked, b, nausea_items, spBlockHeight, spBlockWidth, 15, true, false);
        }
        private void Window_OnSP1CheckBox16Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok16;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_fulltiredness, ref sp1_fulltiredness_chb_ischecked, b, fulltiredness_items, spBlockHeight, spBlockWidth, 16, true, false);
        }
        private void Window_OnSP1CheckBox17Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok17;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_tiredness, ref sp1_tiredness_chb_ischecked, b, tiredness_items, spBlockHeight, spBlockWidth, 17, false, true);
        }
        private void Window_OnSP1CheckBox18Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok18;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_numbness, ref sp1_numbness_chb_ischecked, b, numbness_items, spBlockHeight, spBlockWidth, 18, false, true);
        }
        private void Window_OnSP1CheckBox19Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok19;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_goosebumps, ref sp1_goosebumps_chb_ischecked, b, goosebumps_items, spBlockHeight, spBlockWidth, 19, false, true);
        }
        private void Window_OnSP1CheckBox20Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok20;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_nfto1, ref sp1_nfto1_chb_ischecked, b, nfto1_items, spBlockHeight, spBlockWidth, 20, true, false);
        }
        private void Window_OnSP1CheckBox21Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok21;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_bodydmg, ref sp1_bodydmg_chb_ischecked, b, bodydmg_items, spBlockHeight, spBlockWidth, 21, false, true);
        }
        private void Window_OnSP1CheckBox22Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok22;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_sensivitydmg, ref sp1_sensivitydmg_chb_ischecked, b, sensivitydmg_items, spBlockHeight, spBlockWidth, 22, false, true);
        }
        private void Window_OnSP1CheckBox23Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok23;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_limbsweakness, ref sp1_limbsweakness_chb_ischecked, b, limbsweakness_items, spBlockHeight, spBlockWidth, 23, false, true);
        }
        private void Window_OnSP1CheckBox24Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok24;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_nfto2, ref sp1_nfto2_chb_ischecked, b, nfto2_items, spBlockHeight, spBlockWidth, 24, true, false);
        }
        private void Window_OnSP1CheckBox25Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok25;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_neckdmg, ref sp1_neckdmg_chb_ischecked, b, neckdmg_items, spBlockHeight, spBlockWidth, 25, true, false);
        }
        private void Window_OnSP1CheckBox26Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok26;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_mrt, ref sp1_mrt_chb_ischecked, b, mrt_items, spBlockHeight, spBlockWidth, 26, false, true);
        }
        private void Window_OnSP1CheckBox27Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok27;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_mskt, ref sp1_mskt_chb_ischecked, b, mskt_items, spBlockHeight, spBlockWidth, 27, false, true);
        }
        private void Window_OnSP1CheckBox28Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok28;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_xray, ref sp1_xray_chb_ischecked, b, xray_items, spBlockHeight, spBlockWidth, 28, false, true);
        }
        private void Window_OnSP1CheckBox29Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok29;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_emg, ref sp1_emg_chb_ischecked, b, emg_items, spBlockHeight, spBlockWidth, 29, false, true);
        }
        private void Window_OnSP1CheckBox30Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok30;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_uzi, ref sp1_uzi_chb_ischecked, b, uzi_items, spBlockHeight, spBlockWidth, 30, false, true);
        }
        private void Window_OnSP1CheckBox31Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok31;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_eeg, ref sp1_eeg_chb_ischecked, b, eeg_items, spBlockHeight, spBlockWidth, 31, false, true);
        }
        private void Window_OnSP1CheckBox32Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok32;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_rvg, ref sp1_rvg_chb_ischecked, b, rvg_items, spBlockHeight, spBlockWidth, 32, false, true);
        }
        private void Window_OnSP1CheckBox33Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok33;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_drugoi, ref sp1_drugoi_chb_ischecked, b, drugoi_items, spBlockHeight, spBlockWidth, 33, false, true);
        }
        private void Window_OnSP1CheckBox34Clicked(object sender, RoutedEventArgs e)
        {
            Button b = new Button() { Height = spBlockHeight, Width = spBlockWidth, Content = "Ок", HorizontalAlignment = HorizontalAlignment.Left };
            b.Click += ButtonClick_SP1Ok34;
            Window_FillStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_alcohol, ref sp1_alcohol_chb_ischecked, b, alcohol_items, spBlockHeight, spBlockWidth, 34, false, false);
        }
        // обработчики событий (клавиши):
        private void ButtonClick_LoadData(object sender, RoutedEventArgs e)
        {
            try
            {
                if(dbCon != null && bAdapterConfigured == true)
                {               
                    if(ds.HasChanges())
                    {
                        MessageBoxResult result = MessageBox.Show("Вы не сохранили сделанные вами изменения и при нажатии 'Да' изменения будут отменены, все равно загрузить текущие данные из БД?", "MSAccessClient", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
                        switch (result)
                        {
                            case MessageBoxResult.Yes:
                                //ds.RejectChanges();
                                dbCon.openConnection();
                                ConfigDS();                                
                                dbCon.closeConnection();
                                dg.ItemsSource = ds.Tables[0].DefaultView;
                                MessageBox.Show("Данные из базы загружены.");
                                break;
                            case MessageBoxResult.No:
                                break;
                        }
                    }
                    else
                    {
                        dbCon.openConnection();
                        ConfigDS();
                        dbCon.closeConnection();
                        dg.ItemsSource = ds.Tables[0].DefaultView;
                        MessageBox.Show("Данные из базы загружены.");
                    }                    
                }
                else
                {
                    MessageBox.Show("База данных не выбрана или подключение не настроено.");
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в ButtonClick_LoadData(): " + ex.Message);
            }    
        }
        private void ButtonClick_SelectDB(object sender, RoutedEventArgs e)
        {
            try
            {                
                var fileDialog = new System.Windows.Forms.OpenFileDialog();
                fileDialog.Multiselect = false;

                var result = fileDialog.ShowDialog();
                switch (result)
                {
                    case System.Windows.Forms.DialogResult.OK:
                        string temp = fileDialog.FileName;                        
                        if ((temp.Substring(temp.Length - 6, 6) == ".accdb") || (temp.Substring(temp.Length - 4, 4) == ".mdb"))
                        {
                            int flag = 1;
                            if (dbCon != null && bAdapterConfigured == true)
                            {                                
                                if (ds.HasChanges())
                                {
                                    MessageBoxResult res = MessageBox.Show("Сохранить внесенные изменения в текущую БД перед выбором новой?", "MSAccessClient", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning, MessageBoxResult.Yes);
                                    switch (res)
                                    {
                                        case MessageBoxResult.Yes:
                                            ButtonClick_SaveChanges(null,null);
                                            break;
                                        case MessageBoxResult.No:
                                            break;
                                        case MessageBoxResult.Cancel:
                                            flag = 0;
                                            break;
                                    }
                                }
                            }

                            if(flag == 1)
                            {
                                dbPath = fileDialog.FileName;
                                conString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + @dbPath;
                                dbCon = new DBCon(conString);
                                dbCon.openConnection();

                                if (!dbCon.opened)
                                {
                                    
                                    MessageBox.Show("Не удалось подключиться к базе данных." + dbCon.errormessage);
                                    dbPath = null; conString = null; dbCon = null;
                                }
                                else
                                {
                                    ConfigDSAdapter();
                                    ConfigDS();
                                    dbCon.closeConnection();

                                    dg.CellEditEnding += Window_OnDataGridCellEditEnding;
                                    MessageBox.Show("Подключение к базе данных успешно.");
                                }
                            }                            
                        }
                        else
                        {
                            MessageBox.Show("Программой поддерживаются только базы данных формата .accdb и .mdb. Не удалось подключиться к выбранной базе данных.");
                        }
                        break;

                    case System.Windows.Forms.DialogResult.Cancel:
                        break;
                }
            }
            catch (OleDbException ex)
            {
                dbPath = null; conString = null; dbCon = null;
                MessageBox.Show("Исключение типа OleDbException в ButtonClick_SelectDB(): " + 
                    "Возможно, другой пользователь (или вы) находится в конструкторе таблицы. Закройте Access и попробуйте снова."+ ex.Message);
            }
            catch (Exception ex)
            {
                dbPath = null; conString = null; dbCon = null;
                MessageBox.Show("Исключение в ButtonClick_SelectDB(): " + ex.Message);
            }            
        }
        private void ButtonClick_SaveChanges(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dbCon != null && bAdapterConfigured == true)
                {     
                    DataTable tempDT = ((DataView)dg.ItemsSource).ToTable();
                    dbCon.openConnection();
                    dsAdapter.Fill(tempDT);
                    dsAdapter.Update(ds);
                    ConfigDS();
                    dg.ItemsSource = ds.Tables[0].DefaultView;
                    dbCon.closeConnection();
                    lastChangedRowIndex = -1;

                    MessageBox.Show("Сохранено");
                } 
                else
                {
                    MessageBox.Show("База данных не выбрана или подключение не настроено.");
                }               
            }
            catch(DBConcurrencyException ex)
            {
                // нарушение параллелизма: попытка изменения строк, удаленных другим юзером
                int count = 1;
                int ind = ds.Tables[0].Rows.IndexOf(ex.Row);
                bool existOtherChanges;
                DataRow r;
                
                while(true)
                {                   
                    if(count > ds.Tables[0].Rows.Count)
                    {
                        // защита от зацикливания
                        break;
                    }

                    try
                    {                        
                        r = ds.Tables[0].Rows[ind];
                        /*
                            MessageBox.Show("count = " + count.ToString() + '\n'
                            + "rowName = " + ds.Tables[0].Rows[ind].ItemArray.GetValue(1).ToString() + '\n'
                            + "rowstate = " + ds.Tables[0].Rows[ind].RowState.ToString() + '\n'
                            + "rowerror = " + ds.Tables[0].Rows[ind].RowError+'\n'
                            + "rowhaserrors = " + ds.Tables[0].Rows[ind].HasErrors.ToString());
                        */
                        ds.Tables[0].Rows[ind].RejectChanges();
                        existOtherChanges = ds.HasChanges();
                        dsAdapter.Update(ds);
                        ConfigDS();
                        dg.ItemsSource = ds.Tables[0].DefaultView;
                        dbCon.closeConnection();
                        if(existOtherChanges)
                        {
                            MessageBox.Show("Некоторые измененные вами строки были удалены другим пользователем, работающим с этой базой. Соответствующие изменения отменены, остальные изменения успешно сохранены.");
                        }
                        else
                        {
                            MessageBox.Show("Строки, которые вы модифицировали, были удалены другим пользователем, работающим с этой базой данных. Изменения были отменены.");
                        }
                        break;
                    }
                    catch(DBConcurrencyException innerEx)
                    {
                        ind = ds.Tables[0].Rows.IndexOf(innerEx.Row);
                    }
                    catch(Exception innerEx)
                    {
                        ds.RejectChanges();
                        dsAdapter.Update(ds);
                        ConfigDS();
                        dg.ItemsSource = ds.Tables[0].DefaultView;
                        dbCon.closeConnection();
                        MessageBox.Show("Исключение в ButtonClick_SaveChanges() -> DBConcurrencyException: " + innerEx.GetType().ToString() + " - " + innerEx.Message + "; count = " + count.ToString() + 
                            "." + '\n' + "Необработанное исключение. Изменения не были сохранены");
                        break;
                    }
                    count = count + 1;
                }                             
            }
            
        }
        private void ButtonClick_ClearForm(object sender, RoutedEventArgs e)
        {
            try
            {
                MessageBoxResult res = MessageBox.Show("Очистить форму?", "MSAccessClient", MessageBoxButton.YesNo, MessageBoxImage.None, MessageBoxResult.No);
                switch (res)
                {
                    case MessageBoxResult.Yes:
                        dpk_injurydate.SelectedDate = null;
                        dpk_datemed.SelectedDate = null;
                        cmb_gender.SelectedIndex = -1;
                        cmb_alcohol.SelectedIndex = -1;
                        cmb_regexpert.SelectedIndex = -1;
                        cmb_regtypeexpert.SelectedIndex = -1;
                        cmb_injurytype.SelectedIndex = -1;
                        tbx_age.Clear();
                        tbx_daysbeforemed.Clear();
                        tbx_daystreat000.Clear();  tbx_daystreat001.Clear();  tbx_daystreat002.Clear();                    
                        tbx_fio.Clear();
                        chb_bodydmg000.IsChecked = false; chb_bodydmg001.IsChecked = false; chb_bodydmg002.IsChecked = false; chb_bodydmg003.IsChecked = false; chb_bodydmg004.IsChecked = false;
                        chb_dizziness.IsChecked = false;
                        chb_drugoi000.IsChecked = false; chb_drugoi001.IsChecked = false;
                        chb_eeg000.IsChecked = false; chb_eeg001.IsChecked = false;
                        chb_emg000.IsChecked = false; chb_emg001.IsChecked = false;
                        chb_fulltiredness.IsChecked = false;
                        chb_goosebumps000.IsChecked = false; chb_goosebumps001.IsChecked = false; chb_goosebumps002.IsChecked = false;
                        chb_limbsweakness000.IsChecked = false; chb_limbsweakness001.IsChecked = false;
                        chb_mrt000.IsChecked = false; chb_mrt001.IsChecked = false;
                        chb_mskt000.IsChecked = false; chb_mskt001.IsChecked = false;
                        chb_nausea.IsChecked = false;
                        chb_neckdmg.IsChecked = false;
                        chb_nfto1.IsChecked = false;
                        chb_nfto2.IsChecked = false;
                        chb_numbness000.IsChecked = false; chb_numbness001.IsChecked = false;
                        chb_pain000.IsChecked = false; chb_pain001.IsChecked = false; chb_pain002.IsChecked = false; chb_pain003.IsChecked = false;
                        chb_rvg000.IsChecked = false; chb_rvg001.IsChecked = false;
                        chb_sensivitydmg000.IsChecked = false; chb_sensivitydmg001.IsChecked = false; chb_sensivitydmg002.IsChecked = false; chb_sensivitydmg003.IsChecked = false; chb_sensivitydmg004.IsChecked = false;
                        chb_tiredness000.IsChecked = false; chb_tiredness001.IsChecked = false;
                        chb_uzi000.IsChecked = false; chb_uzi001.IsChecked = false;
                        chb_xray000.IsChecked = false; chb_xray001.IsChecked = false;
                        break;
                    case MessageBoxResult.No:
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в ButtonClick_ClearForm(): " + ex.GetType().ToString() + " " + ex.Message);
            }
        }
        private void ButtonClick_SP1Ok1(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_gender, ref sp1_gender_chb_ischecked, 1, true, false);
        }
        private void ButtonClick_SP1Ok2(object sender, RoutedEventArgs e)
        {           
            Window_ConfirmStackPanelSingleInput(ref sp1, ref sp1_chb, ref sp1_chb_fio, ref sp1_fio_tbx1_text, 2);
        }
        private void ButtonClick_SP1Ok3(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_age, ref sp1_age_cmb1_text, ref sp1_age_tbx1_text, comparison_items1, comparison_signs1, false, true, 3);
        }
        private void ButtonClick_SP1Ok4(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_regexpert, ref sp1_regexpert_chb_ischecked, 4, false, false);
        }
        private void ButtonClick_SP1Ok5(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_regtypeexpert, ref sp1_regtypeexpert_chb_ischecked, 5, false, false);
        }
        private void ButtonClick_SP1Ok6(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_injurydate, ref sp1_injurydate_cmb1_text, ref sp1_injurydate_tbx1_text, comparison_items2, comparison_signs1, true, false, 6);
        }
        private void ButtonClick_SP1Ok7(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_datemed, ref sp1_datemed_cmb1_text, ref sp1_datemed_tbx1_text, comparison_items2, comparison_signs1, true, false, 7);
        }
        private void ButtonClick_SP1Ok8(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_daysbeforemed, ref sp1_daysbeforemed_cmb1_text, ref sp1_daysbeforemed_tbx1_text, comparison_items1, comparison_signs1, false, true, 8);
        }
        private void ButtonClick_SP1Ok9(object sender, RoutedEventArgs e)
        {
           Window_ConfirmStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_daystreat000, ref sp1_daystreat000_cmb1_text, ref sp1_daystreat000_tbx1_text, comparison_items1, comparison_signs1, false, true, 9);
        }
        private void ButtonClick_SP1Ok10(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_daystreat001, ref sp1_daystreat001_cmb1_text, ref sp1_daystreat001_tbx1_text, comparison_items1, comparison_signs1, false, true, 10);
        }
        private void ButtonClick_SP1Ok11(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelScalar(ref sp1, ref sp1_chb, ref sp1_chb_daystreat002, ref sp1_daystreat002_cmb1_text, ref sp1_daystreat002_tbx1_text, comparison_items1, comparison_signs1, false, true, 11);
        }
        private void ButtonClick_SP1Ok12(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_injurytype, ref sp1_injurytype_chb_ischecked, 12, false, false);
        }
        private void ButtonClick_SP1Ok13(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_pain, ref sp1_pain_chb_ischecked, 13, false, true);
        }
        private void ButtonClick_SP1Ok14(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_dizziness, ref sp1_dizziness_chb_ischecked, 14, true, false);
        }
        private void ButtonClick_SP1Ok15(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_nausea, ref sp1_nausea_chb_ischecked, 15, true, false);
        }
        private void ButtonClick_SP1Ok16(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_fulltiredness, ref sp1_fulltiredness_chb_ischecked, 16, true, false);
        }
        private void ButtonClick_SP1Ok17(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_tiredness, ref sp1_tiredness_chb_ischecked, 17, false, true);
        }
        private void ButtonClick_SP1Ok18(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_numbness, ref sp1_numbness_chb_ischecked, 18, false, true);
        }
        private void ButtonClick_SP1Ok19(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_goosebumps, ref sp1_goosebumps_chb_ischecked, 19, false, true);
        }
        private void ButtonClick_SP1Ok20(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_nfto1, ref sp1_nfto1_chb_ischecked, 20, true, false);
        }
        private void ButtonClick_SP1Ok21(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_bodydmg, ref sp1_bodydmg_chb_ischecked, 21, false, true);
        }
        private void ButtonClick_SP1Ok22(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_sensivitydmg, ref sp1_sensivitydmg_chb_ischecked, 22, false, true);
        }
        private void ButtonClick_SP1Ok23(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_limbsweakness, ref sp1_limbsweakness_chb_ischecked, 23, false, true);
        }
        private void ButtonClick_SP1Ok24(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_nfto2, ref sp1_nfto2_chb_ischecked, 24, true, false);
        }
        private void ButtonClick_SP1Ok25(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_neckdmg, ref sp1_neckdmg_chb_ischecked, 25, true, false);
        }
        private void ButtonClick_SP1Ok26(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_mrt, ref sp1_mrt_chb_ischecked, 26, false, true);
        }
        private void ButtonClick_SP1Ok27(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_mskt, ref sp1_mskt_chb_ischecked, 27, false, true);
        }
        private void ButtonClick_SP1Ok28(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_xray, ref sp1_xray_chb_ischecked, 28, false, true);
        }
        private void ButtonClick_SP1Ok29(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_emg, ref sp1_emg_chb_ischecked, 29, false, true);
        }
        private void ButtonClick_SP1Ok30(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_uzi, ref sp1_uzi_chb_ischecked, 30, false, true);
        }
        private void ButtonClick_SP1Ok31(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_eeg, ref sp1_eeg_chb_ischecked, 31, false, true);
        }
        private void ButtonClick_SP1Ok32(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_rvg, ref sp1_rvg_chb_ischecked, 32, false, true);
        }
        private void ButtonClick_SP1Ok33(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_drugoi, ref sp1_drugoi_chb_ischecked, 33, false, true);
        }
        private void ButtonClick_SP1Ok34(object sender, RoutedEventArgs e)
        {
            Window_ConfirmStackPanelCategorical(ref sp1, ref sp1_chb, ref sp1_chb_alcohol, ref sp1_alcohol_chb_ischecked, 34, false, false);
        }
        private void ButtonClick_ShowFormParams(object sender, RoutedEventArgs e)
        {
            try
            {
                string res = "";

                res = res + " ------------ " + "Общие сведения" + " ------------ " + '\n';
                /*[cat]*/
                res = res + ">>" + "Пол: "; for (int i = 0; i < 2; i++) { res = res + gender_items[i] + ": " + sp1_gender_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                /*[sinp]*/
                res = res + ">>" + "ФИО: "; if (sp1_fio_tbx1_text == "") { res = res + "-" + '\n'; } else { res = res + sp1_fio_tbx1_text + '\n'; };
                /*[scalar]*/
                res = res + ">>" + "Возраст: "; if (sp1_age_cmb1_text == "") { res = res + "-" + '\n'; } else { res = res + sp1_age_cmb1_text + " " + sp1_age_tbx1_text + '\n'; };
                res = res + ">>" + "Рег. вып. эксп: "; for (int i = 0; i < 7; i++) { res = res + regexpert_items[i] + ": " + sp1_regexpert_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "Тип нас. п: "; for (int i = 0; i < 5; i++) { res = res + regtypeexpert_items[i] + ": " + sp1_regtypeexpert_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';

                res = res + " ------------ " + "Анамнез травмирования" + " ------------ " + '\n';
                res = res + ">>" + "Дата травмы: "; if (sp1_injurydate_cmb1_text == "") { res = res + "-" + '\n'; } else { res = res + sp1_injurydate_cmb1_text + " " + sp1_injurydate_tbx1_text + '\n'; };
                res = res + ">>" + "Дата обр. мед.: "; if (sp1_datemed_cmb1_text == "") { res = res + "-" + '\n'; } else { res = res + sp1_datemed_cmb1_text + " " + sp1_datemed_tbx1_text + '\n'; };
                res = res + ">>" + "Срок до обр: "; if (sp1_daysbeforemed_cmb1_text == "") { res = res + "-" + '\n'; } else { res = res + sp1_daysbeforemed_cmb1_text + " " + sp1_daysbeforemed_tbx1_text + '\n'; };
                res = res + ">>" + "Длит. лечения стационар: "; if (sp1_daystreat000_cmb1_text == "") { res = res + "-" + '\n'; } else { res = res + sp1_daystreat000_cmb1_text + " " + sp1_daystreat000_tbx1_text + '\n'; };
                res = res + ">>" + "Длит. лечения амбулаторно: "; if (sp1_daystreat001_cmb1_text == "") { res = res + "-" + '\n'; } else { res = res + sp1_daystreat001_cmb1_text + " " + sp1_daystreat001_tbx1_text + '\n'; };
                res = res + ">>" + "Длит. лечения реабилитация: "; if (sp1_daystreat002_cmb1_text == "") { res = res + "-" + '\n'; } else { res = res + sp1_daystreat002_cmb1_text + " " + sp1_daystreat002_tbx1_text + '\n'; };

                res = res + ">>" + "Вид травмы: "; for (int i = 0; i < 7; i++) { res = res + injurytype_items[i] + ": " + sp1_injurytype_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';

                res = res + " ------------ " + "Субъективные данные" + " ------------ " + '\n';
                res = res + ">>" + "Боль: "; for (int i = 0; i < 4; i++) { res = res + pain_items[i] + ": " + sp1_pain_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "Головокружение: "; for (int i = 0; i < 2; i++) { res = res + dizziness_items[i] + ": " + sp1_dizziness_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "Тошнота: "; for (int i = 0; i < 2; i++) { res = res + nausea_items[i] + ": " + sp1_nausea_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "Слабость общ.: "; for (int i = 0; i < 2; i++) { res = res + fulltiredness_items[i] + ": " + sp1_fulltiredness_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "Слабость: "; for (int i = 0; i < 2; i++) { res = res + tiredness_items[i] + ": " + sp1_tiredness_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "Онемение: "; for (int i = 0; i < 2; i++) { res = res + numbness_items[i] + ": " + sp1_numbness_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "Чув. мурашек: "; for (int i = 0; i < 2; i++) { res = res + goosebumps_items[i] + ": " + sp1_goosebumps_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "НФТО (суб): "; for (int i = 0; i < 2; i++) { res = res + nfto1_items[i] + ": " + sp1_nfto1_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';

                res = res + " ------------ " + "Объективные данные" + " ------------ " + '\n';
                res = res + ">>" + "Нар. тел. повр: "; for (int i = 0; i < 5; i++) { res = res + bodydmg_items[i] + ": " + sp1_bodydmg_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "Наруш. чувст: "; for (int i = 0; i < 5; i++) { res = res + sensivitydmg_items[i] + ": " + sp1_sensivitydmg_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "Ум. силы в кон: "; for (int i = 0; i < 2; i++) { res = res + limbsweakness_items[i] + ": " + sp1_limbsweakness_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "НФТО (об): "; for (int i = 0; i < 2; i++) { res = res + nfto2_items[i] + ": " + sp1_nfto2_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "Мыш. деф. шеи: "; for (int i = 0; i < 2; i++) { res = res + neckdmg_items[i] + ": " + sp1_neckdmg_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';

                res = res + " ------------ " + "Инструментальные методы иссл." + " ------------ " + '\n';
                res = res + ">>" + "МРТ: "; for (int i = 0; i < 2; i++) { res = res + mrt_items[i] + ": " + sp1_mrt_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "МСКТ: "; for (int i = 0; i < 2; i++) { res = res + mskt_items[i] + ": " + sp1_mskt_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "МСКТ: "; for (int i = 0; i < 2; i++) { res = res + xray_items[i] + ": " + sp1_xray_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "МСКТ: "; for (int i = 0; i < 2; i++) { res = res + emg_items[i] + ": " + sp1_emg_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "МСКТ: "; for (int i = 0; i < 2; i++) { res = res + uzi_items[i] + ": " + sp1_uzi_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "МСКТ: "; for (int i = 0; i < 2; i++) { res = res + eeg_items[i] + ": " + sp1_eeg_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "МСКТ: "; for (int i = 0; i < 2; i++) { res = res + rvg_items[i] + ": " + sp1_rvg_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                res = res + ">>" + "МСКТ: "; for (int i = 0; i < 2; i++) { res = res + drugoi_items[i] + ": " + sp1_drugoi_chb_ischecked[i].ToString() + ", "; }; res = res + '\n';
                MessageBox.Show(res);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в ButtonClick_Add(): " + ex.GetType().ToString() + " " + ex.Message);
            }


        }
        private void ButtonClick_SelectQuery(object sender, RoutedEventArgs e)
        {
            try
            {
                bool anychecked = false;
                int l;

                if (dbCon != null && conString != null && bAdapterConfigured == true)
                {
                    DataTable resultDT = new DataTable();
                    string qstr = "SELECT * FROM " + tableName + " WHERE ";
                    OleDbCommand oledbcommand = new OleDbCommand("", dbCon.getConnection());

                    qstr = qstr + BuildCommandStringHelperMultiCatCase(ref sp1_chb_gender, ref anychecked, "Пол", sp1_gender_chb_ischecked, gender_items, "AND", true, false);
                    qstr = qstr + BuildCommandStringHelperSingleCase(ref oledbcommand, ref sp1_chb_fio, ref anychecked, "ФамилияИО", 1, sp1_fio_tbx1_text, "LIKE");
                    qstr = qstr + BuildCommandStringHelperSingleCase(ref oledbcommand, ref sp1_chb_age, ref anychecked, "Возраст", 2, sp1_age_tbx1_text, sp1_age_cmb1_text);
                    qstr = qstr + BuildCommandStringHelperMultiCatCase(ref sp1_chb_regexpert, ref anychecked, "РегионВыпЭксп", sp1_regexpert_chb_ischecked, regexpert_items, "OR", false, true);
                    qstr = qstr + BuildCommandStringHelperMultiCatCase(ref sp1_chb_regtypeexpert, ref anychecked, "ТипНасПунктаТравм", sp1_regtypeexpert_chb_ischecked, regtypeexpert_items, "OR", false, false);
                    qstr = qstr + BuildCommandStringHelperSingleCase(ref oledbcommand, ref sp1_chb_injurydate, ref anychecked, "ДатаТравмы", 3, sp1_injurydate_tbx1_text, sp1_injurydate_cmb1_text);
                    qstr = qstr + BuildCommandStringHelperSingleCase(ref oledbcommand, ref sp1_chb_datemed, ref anychecked, "ДатаОбрЗаПом", 3, sp1_datemed_tbx1_text, sp1_datemed_cmb1_text);
                    qstr = qstr + BuildCommandStringHelperSingleCase(ref oledbcommand, ref sp1_chb_daysbeforemed, ref anychecked, "СрокСМомТравмДоОбрЗаПом", 2, sp1_daysbeforemed_tbx1_text, sp1_daysbeforemed_cmb1_text);
                    qstr = qstr + BuildCommandStringHelperSingleCase(ref oledbcommand, ref sp1_chb_daystreat000, ref anychecked, "ДлитЛечСтационар", 2, sp1_daystreat000_tbx1_text, sp1_daystreat000_cmb1_text);
                    qstr = qstr + BuildCommandStringHelperSingleCase(ref oledbcommand, ref sp1_chb_daystreat001, ref anychecked, "ДлитЛечАмбулаторно", 2, sp1_daystreat001_tbx1_text, sp1_daystreat001_cmb1_text);
                    qstr = qstr + BuildCommandStringHelperSingleCase(ref oledbcommand, ref sp1_chb_daystreat002, ref anychecked, "ДлитЛечРеабилитация", 2, sp1_daystreat002_tbx1_text, sp1_daystreat002_cmb1_text);
                    qstr = qstr + BuildCommandStringHelperMultiCatCase(ref sp1_chb_injurytype, ref anychecked, "ВидТравмы", sp1_injurytype_chb_ischecked, injurytype_items, "OR", false, false);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_pain, ref anychecked, pain_columns, sp1_pain_chb_ischecked, sp1_cat_operators[13]);
                    qstr = qstr + BuildCommandStringHelperSingleBoolCase(ref sp1_chb_dizziness, ref anychecked, "Головокружение", sp1_dizziness_chb_ischecked);
                    qstr = qstr + BuildCommandStringHelperSingleBoolCase(ref sp1_chb_nausea, ref anychecked, "ТошнотаРвота", sp1_nausea_chb_ischecked);
                    qstr = qstr + BuildCommandStringHelperSingleBoolCase(ref sp1_chb_fulltiredness, ref anychecked, "СлабостьОбщая", sp1_fulltiredness_chb_ischecked);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_tiredness, ref anychecked, tiredness_columns, sp1_tiredness_chb_ischecked, sp1_cat_operators[17]);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_numbness, ref anychecked, numbness_columns, sp1_numbness_chb_ischecked, sp1_cat_operators[18]);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_goosebumps, ref anychecked, goosebumps_columns, sp1_goosebumps_chb_ischecked, sp1_cat_operators[19]);
                    qstr = qstr + BuildCommandStringHelperSingleBoolCase(ref sp1_chb_nfto1, ref anychecked, "НФТО1", sp1_nfto1_chb_ischecked);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_bodydmg, ref anychecked, bodydmg_columns, sp1_bodydmg_chb_ischecked, sp1_cat_operators[21]);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_sensivitydmg, ref anychecked, sensivitydmg_columns, sp1_sensivitydmg_chb_ischecked, sp1_cat_operators[22]);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_limbsweakness, ref anychecked, limbsweakness_columns, sp1_limbsweakness_chb_ischecked, sp1_cat_operators[23]);
                    qstr = qstr + BuildCommandStringHelperSingleBoolCase(ref sp1_chb_nfto2, ref anychecked, "НФТО2", sp1_nfto2_chb_ischecked);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_mrt, ref anychecked, mrt_columns, sp1_mrt_chb_ischecked, sp1_cat_operators[26]);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_mskt, ref anychecked, mskt_columns, sp1_mskt_chb_ischecked, sp1_cat_operators[27]);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_xray, ref anychecked, xray_columns, sp1_xray_chb_ischecked, sp1_cat_operators[28]);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_emg, ref anychecked, emg_columns, sp1_emg_chb_ischecked, sp1_cat_operators[29]);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_uzi, ref anychecked, uzi_columns, sp1_uzi_chb_ischecked, sp1_cat_operators[30]);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_eeg, ref anychecked, eeg_columns, sp1_eeg_chb_ischecked, sp1_cat_operators[31]);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_rvg, ref anychecked, rvg_columns, sp1_rvg_chb_ischecked, sp1_cat_operators[32]);
                    qstr = qstr + BuildCommandStringHelperMultiBoolCase(ref sp1_chb_drugoi, ref anychecked, drugoi_columns, sp1_drugoi_chb_ischecked, sp1_cat_operators[33]);
                    qstr = qstr + BuildCommandStringHelperMultiCatCase(ref sp1_chb_alcohol, ref anychecked, "Алкоголь", sp1_alcohol_chb_ischecked, alcohol_items, "OR", false, false);

                    l = qstr.Length;
                    if (qstr.Substring(l - 3, 3) == "AND")
                    {
                        qstr = qstr.Substring(0, l - 4) + ";";
                    }

                    if (anychecked)
                    {
                        qstr = qstr.Replace("= 'Нет данных'", "IS NULL");
                        MessageBox.Show(qstr);

                        dbCon.openConnection();
                        oledbcommand.CommandText = qstr;
                        using (OleDbDataReader oledbdatareader = oledbcommand.ExecuteReader())
                        {
                            resultDT.Load(oledbdatareader);
                            resultDT.Select();
                            
                            dg2.ItemsSource = resultDT.DefaultView;
                            oledbdrparamsdatecounter = 1;
                        }
                        dbCon.closeConnection();

                        tbx_result.Text = resultDT.Rows.Count.ToString();
                    }
                    else
                    {
                        MessageBox.Show("Вы не выбрали параметров запроса");
                    }
                }
                else
                {
                    MessageBox.Show("База данных не выбрана или подключение не настроено.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в ButtonClick_SelectQuery(): " + ex.GetType().ToString() + " " + ex.Message);
            }
        }
        private void ButtonClick_Add(object sender, RoutedEventArgs e)
        {
            int temp;
            try
            {
                if (dbCon != null && bAdapterConfigured == true)
                {
                    DataRow dr = ds.Tables[0].NewRow();

                    dr[1] = tbx_fio.Text;

                    if (tbx_age.Text != "")
                    {
                        if (int.TryParse(tbx_age.Text, out temp))
                        {
                            if (temp >= 0)
                            {
                                dr[2] = temp;

                            }
                            else
                            {
                                throw new InvalidDataException("Возраст не может быть отрицательным.");
                            }
                        }
                        else
                        {
                            throw new InvalidDataException("В графу 'Возраст' введена строка, которая не является целым числом. Оставьте строку пустой, если возраст не нужно указывать.");
                        }
                    }

                    if (cmb_gender.Text != "") { dr[3] = cmb_gender.Text; } else { dr[3] = DBNull.Value; }

                    if (cmb_regexpert.Text != "") { dr[4] = int.Parse(cmb_regexpert.Text.Substring(0, 1)); } else { dr[4] = DBNull.Value; }

                    if (cmb_regtypeexpert.Text != "") { dr[5] = cmb_regtypeexpert.Text; } else { dr[5] = DBNull.Value; }

                    if (dpk_injurydate.SelectedDate.HasValue) { dr[6] = dpk_injurydate.SelectedDate.Value;} else { dr[6] = DBNull.Value;}

                    if (dpk_datemed.SelectedDate.HasValue) { dr[7] = dpk_datemed.SelectedDate.Value; } else { dr[7] = DBNull.Value; }

                    if (tbx_daysbeforemed.Text != "")
                    {
                        if (int.TryParse(tbx_daysbeforemed.Text, out temp))
                        {
                            if (temp >= 0)
                            {
                                dr[8] = temp;
                            }
                            else
                            {
                                throw new InvalidDataException("Срок с момента травмы до обр. за медпомощью не может быть отрицательным.");
                            }
                        }
                        else
                        {
                            throw new InvalidDataException("В графу 'Срок с мом травмы до обр. за медпомощью' введена строка, которая не является целым числом. Оставьте строку пустой, если возраст не нужно указывать.");
                        }
                    }

                    if (tbx_daystreat000.Text != "")
                    {
                        if (int.TryParse(tbx_daystreat000.Text, out temp))
                        {
                            if (temp >= 0)
                            {
                                dr[9] = temp;
                            }
                            else
                            {
                                throw new InvalidDataException("Длительность лечения (стационар) не может быть отрицательным.");
                            }
                        }
                        else
                        {
                            throw new InvalidDataException("В графу 'Длительность лечения (стационар)' введена строка, которая не является целым числом. Оставьте строку пустой, если не нужно указывать.");
                        }
                    }

                    if (tbx_daystreat001.Text != "")
                    {
                        if (int.TryParse(tbx_daystreat001.Text, out temp))
                        {
                            if (temp >= 0)
                            {
                                dr[10] = temp;
                            }
                            else
                            {
                                throw new InvalidDataException("Длительность лечения (реабилитация) не может быть отрицательным.");
                            }
                        }
                        else
                        {
                            throw new InvalidDataException("В графу 'Длительность лечения (реабилитация)' введена строка, которая не является целым числом. Оставьте строку пустой, если не нужно указывать.");
                        }
                    }
                    if (tbx_daystreat002.Text != "")
                    {
                        if (int.TryParse(tbx_daystreat002.Text, out temp))
                        {
                            if (temp >= 0)
                            {
                                dr[11] = temp;
                            }
                            else
                            {
                                throw new InvalidDataException("Длительность лечения (амбулаторно) не может быть отрицательным.");
                            }
                        }
                        else
                        {
                            throw new InvalidDataException("В графу 'Длительность лечения (стационар)' введена строка, которая не является целым числом. Оставьте строку пустой, если не нужно указывать.");
                        }
                    }

                    if (cmb_injurytype.Text != "") { dr[12] = cmb_injurytype.Text; } else { dr[12] = DBNull.Value; }

                    dr[13] = chb_pain000.IsChecked.HasValue ? chb_pain000.IsChecked.Value : false;
                    dr[14] = chb_pain001.IsChecked.HasValue ? chb_pain001.IsChecked.Value : false;
                    dr[15] = chb_pain002.IsChecked.HasValue ? chb_pain002.IsChecked.Value : false;
                    dr[16] = chb_pain003.IsChecked.HasValue ? chb_pain003.IsChecked.Value : false;

                    dr[17] = chb_dizziness.IsChecked.HasValue ? chb_dizziness.IsChecked.Value : false;
                    dr[18] = chb_nausea.IsChecked.HasValue ? chb_nausea.IsChecked.Value : false;
                    dr[19] = chb_fulltiredness.IsChecked.HasValue ? chb_fulltiredness.IsChecked.Value : false;

                    dr[20] = chb_tiredness000.IsChecked.HasValue ? chb_tiredness000.IsChecked.Value : false;
                    dr[21] = chb_tiredness001.IsChecked.HasValue ? chb_tiredness001.IsChecked.Value : false;

                    dr[22] = chb_numbness000.IsChecked.HasValue ? chb_numbness000.IsChecked.Value : false;
                    dr[23] = chb_numbness001.IsChecked.HasValue ? chb_numbness001.IsChecked.Value : false;

                    dr[24] = chb_goosebumps000.IsChecked.HasValue ? chb_goosebumps000.IsChecked.Value : false;
                    dr[25] = chb_goosebumps001.IsChecked.HasValue ? chb_goosebumps001.IsChecked.Value : false;
                    dr[26] = chb_goosebumps002.IsChecked.HasValue ? chb_goosebumps002.IsChecked.Value : false;

                    dr[27] = chb_nfto1.IsChecked.HasValue ? chb_nfto1.IsChecked.Value : false;

                    dr[28] = chb_bodydmg000.IsChecked.HasValue ? chb_bodydmg000.IsChecked.Value : false;
                    dr[29] = chb_bodydmg001.IsChecked.HasValue ? chb_bodydmg001.IsChecked.Value : false;
                    dr[30] = chb_bodydmg002.IsChecked.HasValue ? chb_bodydmg002.IsChecked.Value : false;
                    dr[31] = chb_bodydmg003.IsChecked.HasValue ? chb_bodydmg003.IsChecked.Value : false;
                    dr[32] = chb_bodydmg004.IsChecked.HasValue ? chb_bodydmg004.IsChecked.Value : false;

                    dr[33] = chb_sensivitydmg000.IsChecked.HasValue ? chb_sensivitydmg000.IsChecked.Value : false;
                    dr[34] = chb_sensivitydmg001.IsChecked.HasValue ? chb_sensivitydmg001.IsChecked.Value : false;
                    dr[35] = chb_sensivitydmg002.IsChecked.HasValue ? chb_sensivitydmg002.IsChecked.Value : false;
                    dr[36] = chb_sensivitydmg003.IsChecked.HasValue ? chb_sensivitydmg003.IsChecked.Value : false;
                    dr[37] = chb_sensivitydmg004.IsChecked.HasValue ? chb_sensivitydmg004.IsChecked.Value : false;

                    dr[38] = chb_limbsweakness000.IsChecked.HasValue ? chb_limbsweakness000.IsChecked.Value : false;
                    dr[39] = chb_limbsweakness001.IsChecked.HasValue ? chb_limbsweakness001.IsChecked.Value : false;

                    dr[40] = chb_nfto2.IsChecked.HasValue ? chb_nfto2.IsChecked.Value : false;
                    dr[41] = chb_neckdmg.IsChecked.HasValue ? chb_neckdmg.IsChecked.Value : false;

                    dr[42] = chb_mrt000.IsChecked.HasValue ? chb_mrt000.IsChecked.Value : false;
                    dr[43] = chb_mrt001.IsChecked.HasValue ? chb_mrt001.IsChecked.Value : false;

                    dr[44] = chb_mskt000.IsChecked.HasValue ? chb_mskt000.IsChecked.Value : false;
                    dr[45] = chb_mskt001.IsChecked.HasValue ? chb_mskt001.IsChecked.Value : false;

                    dr[46] = chb_xray000.IsChecked.HasValue ? chb_xray000.IsChecked.Value : false;
                    dr[47] = chb_xray001.IsChecked.HasValue ? chb_xray001.IsChecked.Value : false;

                    dr[48] = chb_emg000.IsChecked.HasValue ? chb_emg000.IsChecked.Value : false;
                    dr[49] = chb_emg001.IsChecked.HasValue ? chb_emg001.IsChecked.Value : false;

                    dr[50] = chb_uzi000.IsChecked.HasValue ? chb_uzi000.IsChecked.Value : false;
                    dr[51] = chb_uzi001.IsChecked.HasValue ? chb_uzi001.IsChecked.Value : false;

                    dr[52] = chb_eeg000.IsChecked.HasValue ? chb_eeg000.IsChecked.Value : false;
                    dr[53] = chb_eeg001.IsChecked.HasValue ? chb_eeg001.IsChecked.Value : false;

                    dr[54] = chb_rvg000.IsChecked.HasValue ? chb_rvg000.IsChecked.Value : false;
                    dr[55] = chb_rvg001.IsChecked.HasValue ? chb_rvg001.IsChecked.Value : false;

                    dr[56] = chb_drugoi000.IsChecked.HasValue ? chb_drugoi000.IsChecked.Value : false;
                    dr[57] = chb_drugoi001.IsChecked.HasValue ? chb_drugoi001.IsChecked.Value : false;

                    if (cmb_alcohol.Text != "" && cmb_alcohol.Text != "Нет данных")
                    {
                        dr[58] = cmb_alcohol.Text;
                    }
                    else
                    {
                        dr[58] = DBNull.Value;
                    }


                    ds.Tables[0].Rows.Add(dr);
                    MessageBox.Show("Строка добавлена.");

                    //ClearForm();
                }
                else
                {
                    MessageBox.Show("База данных не выбрана или подключение не настроено.");
                }
            }
            catch (InvalidDataException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Исключение в ButtonClick_Add(): " + ex.GetType().ToString() + " " + ex.Message);
            }


        }
    }
}

