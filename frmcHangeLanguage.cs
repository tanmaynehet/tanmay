using System;
using System.Collections.Generic; 
using System.Drawing;
using System.Collections;
using System.Globalization;
using System.ComponentModel;
using System.Windows.Forms;
using INI;

namespace cost_management
{
    public partial class frmcHangeLanguage : Form
    {
        IniFile ini = new IniFile();
        public frmcHangeLanguage()
        {
            InitializeComponent();
            LanguageCollector lc = new LanguageCollector();
            //LanguageCollector lc = new LanguageCollector(CultureInfo.CurrentUICulture);
            int currentLanguage = 0;
            CultureInfoDisplayItem[] lis = lc.GetLanguages(System.Globalization.LanguageCollector.LanguageNameDisplay.NativeName, out currentLanguage);
            comboBox.Items.AddRange(lis);
            comboBox.SelectedIndex = currentLanguage;
        }

        private void frmcHangeLanguage_Load(object sender, EventArgs e)
        {
            if (!ini.IniReadValue("CultureInfo", "Language").ToString().Equals(""))
            {
                comboBox.SelectedIndex = Convert.ToInt32(ini.IniReadValue("CultureInfo", "LanguageIndex").ToString());
                System.Globalization.FormLanguageSwitchSingleton.Instance.ChangeLanguage(this, CultureInfo);
            }
        }
        public CultureInfo CultureInfo
        {
            get { return ((CultureInfoDisplayItem)comboBox.SelectedItem).CultureInfo; }
        }

        public bool ChangeCurrentThreadLanguage
        {
            get { return false; }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            ini.IniWriteValue("CultureInfo", "LanguageIndex",comboBox.SelectedIndex.ToString());
            ini.IniWriteValue("CultureInfo", "Language", CultureInfo.ToString());
        }
    }
}
