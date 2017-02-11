using Autodesk.Cad.Crushner.Core;
using Autodesk.Cad.Crushner.Settings;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace Dawager
{
    /// <summary>
    /// Главная форма приложения
    /// </summary>
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();

            _prevLevel = LEVEL.UNKNOWN;
        }

        private enum LEVEL { UNKNOWN = -1, BLOCK, ENTITY }

        private LEVEL _prevLevel;

        private void treeViewBlockDefinition_AfterSelect(object sender, TreeViewEventArgs e)
        {
            int iColumn = -1
                , iRow = -1;
            string blockName = string.Empty;
            EntityParser.ProxyEntity pEntity;

            // строки удаляем всегда
            m_dgvBlockDefinitionProperties.Rows.Clear();
            // столбцы удаляем, если изменился выбранный тип (уровень)
            if (!(_prevLevel == (LEVEL)e.Node.Level)) {
                m_dgvBlockDefinitionProperties.Columns.Clear();
                // добавляем столбцы
                switch ((LEVEL)e.Node.Level) {
                    case LEVEL.BLOCK:
                        m_dgvBlockDefinitionProperties.Columns.AddRange(new DataGridViewColumn[] {
                            new DataGridViewTextBoxColumn()
                            //, new DataGridViewTextBoxColumn()
                            , new DataGridViewTextBoxColumn()
                        });

                        m_dgvBlockDefinitionProperties.ColumnHeadersVisible = true;
                        m_dgvBlockDefinitionProperties.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToDisplayedHeaders | DataGridViewRowHeadersWidthSizeMode.DisableResizing;

                        iColumn = 0;
                        m_dgvBlockDefinitionProperties.Columns[iColumn].HeaderText = @"Имя";
                        //m_dgvBlockDefinitionProperties.Columns[iColumn].Width = 60;
                        m_dgvBlockDefinitionProperties.Columns[iColumn].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        //m_dgvBlockDefinitionProperties.Columns[iColumn].Frozen = true;
                        //iColumn++;
                        //m_dgvBlockDefinitionProperties.Columns[iColumn].HeaderText = @"Тип";
                        //m_dgvBlockDefinitionProperties.Columns[iColumn].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        iColumn++;
                        m_dgvBlockDefinitionProperties.Columns[iColumn].HeaderText = @"Отображать";
                        m_dgvBlockDefinitionProperties.Columns[iColumn].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        break;
                    case LEVEL.ENTITY:
                        m_dgvBlockDefinitionProperties.Columns.AddRange(new DataGridViewColumn[] {
                            new DataGridViewTextBoxColumn()
                            //, new DataGridViewTextBoxColumn()
                        });

                        m_dgvBlockDefinitionProperties.ColumnHeadersVisible = true;
                        m_dgvBlockDefinitionProperties.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToDisplayedHeaders | DataGridViewRowHeadersWidthSizeMode.DisableResizing;

                        iColumn = 0;
                        m_dgvBlockDefinitionProperties.Columns[iColumn].HeaderText = @"Значение";
                        m_dgvBlockDefinitionProperties.Columns[iColumn].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        //iColumn++;
                        //m_dgvBlockDefinitionProperties.Columns[iColumn].HeaderText = @"№ столбца";
                        //m_dgvBlockDefinitionProperties.Columns[iColumn].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        break;
                    default:
                        break;
                }
            } else
                ;
            // добавляем строки
            switch ((LEVEL)e.Node.Level) {
                case LEVEL.BLOCK:
                    blockName = (string)e.Node.Tag;

                    foreach (EntityParser.ProxyEntity entity in MSExcel.s_dictBlock[blockName].m_dictEntityParser.Values) {
                        iRow = m_dgvBlockDefinitionProperties.Rows.Add(new object[] { entity.m_name, true });
                        m_dgvBlockDefinitionProperties.Rows[iRow].HeaderCell.Value = entity.m_command.ToString();
                    }
                    break;
                case LEVEL.ENTITY:
                    blockName = (string)e.Node.Parent.Tag;
                    pEntity = MSExcel.s_dictBlock[blockName].m_dictEntityParser[(KEY_ENTITY)e.Node.Tag];

                    foreach (EntityParser.ProxyEntity.Property property in pEntity.Properties) {
                        iRow = m_dgvBlockDefinitionProperties.Rows.Add(new object[] { property.Value });
                        m_dgvBlockDefinitionProperties.Rows[iRow].HeaderCell.Value = property.Index.ToString();
                    }
                    break;
                default:
                    break;
            }

            _prevLevel = (LEVEL)e.Node.Level;
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
            TreeNode nodeBlock
                , nodeEntity;
            int iRow = -1;

            try {
                MSExcel.Clear();
                MSExcel.Import(m_listBoxFileSettings.SelectedItem.ToString(), MSExcel.FORMAT.HEAP);

                m_treeViewBlockDefinition.Nodes.Clear();

                foreach (string nameBlock in MSExcel.s_dictBlock.Keys) {
                    nodeBlock = m_treeViewBlockDefinition.Nodes.Add(nameBlock);

                    nodeBlock.Tag = nameBlock;
                    nodeBlock.Checked = true;

                    foreach (KeyValuePair<KEY_ENTITY, EntityParser.ProxyEntity> pair in MSExcel.s_dictBlock[nameBlock].m_dictEntityParser) {
                        nodeEntity = nodeBlock.Nodes.Add(pair.Value.m_name);

                        nodeEntity.Tag = pair.Key;
                        nodeEntity.Checked = true;
                    }

                    foreach(MSExcel.BLOCK.PLACEMENT placement in MSExcel.s_dictBlock[nameBlock].m_ListReference) {
                        iRow = m_clbBlockReferences.Items.Add(string.Format(@"{0} ({1})"
                            , nameBlock, placement.ToString()), true);
                    }
                }

                //m_treeViewBlockDefinition.TopNode.Expand();
                m_treeViewBlockDefinition.SelectedNode = m_treeViewBlockDefinition.Nodes[0];
            } catch (Exception e) {
                Logging.ExceptionCaller(MethodBase.GetCurrentMethod(), e);
            }
        }

        private void buttonBlockDefinitionChecked_Click(object sender, EventArgs e)
        {

        }

        private void buttonBlockDefinitionUnchecked_Click(object sender, EventArgs e)
        {

        }
    }
}
