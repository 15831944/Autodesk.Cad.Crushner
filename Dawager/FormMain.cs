using Autodesk.Cad.Crushner.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace Dawager
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void выходToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void m_buttonExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            string curPath = Environment.CurrentDirectory;

            List<string>listFullPathFilesettings =
                Directory.GetFiles(curPath, @"*.xls", SearchOption.TopDirectoryOnly).ToList();

            listFullPathFilesettings.ForEach(fullPathFileSettings => {
                m_listBoxFileSettings.Items.Add(Path.GetFileName(fullPathFileSettings));
            });

            if (m_listBoxFileSettings.Items.Count > 0)
                m_listBoxFileSettings.SelectedIndex = 0;
            else
                ;
        }

        private void listBoxFileSettings_SelectedIndexChanged(object sender, EventArgs ev)
        {
            try {
                MSExcel.Clear();
                MSExcel.Import(m_listBoxFileSettings.SelectedItem.ToString(), MSExcel.FORMAT.HEAP);
            } catch (Exception e) {
                Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e);
            }
        }
    }
}
