using SolidEdgeCommunity;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using SolidEdgePart;
using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Samsonov25
{
    public partial class Form1 : Form
    {
        private SolidEdgeFramework.Application application;
        private SolidEdgeFramework.SolidEdgeDocument document;
        private SolidEdgeFramework.Variables variables;
        private SolidEdgeFramework.VariableList dimensions;
        private Panel panel;
        private DataGridView dataGridView;

        public Form1()
        {
            InitializeComponent();
            try
            {
                OleMessageFilter.Register();
                application = SolidEdgeUtils.Connect(true);

                if (application.ActiveDocumentType != DocumentTypeConstants.igPartDocument)
                {
                    return;
                }

                document = (SolidEdgeDocument)application.ActiveDocument;
                variables = (Variables)document.Variables;

                dimensions = (VariableList)variables.Query(
                    pFindCriterium: "*",
                    NamedBy: SolidEdgeConstants.VariableNameBy.seVariableNameByBoth,
                    VarType: SolidEdgeConstants.VariableVarType.SeVariableVarTypeDimension
                );

                panel = new Panel();
                panel.AutoScroll = true;
                panel.Location = new System.Drawing.Point(0, 0);
                panel.Size = new System.Drawing.Size(580, 450);
                Controls.Add(panel);

                dataGridView = new DataGridView();
                dataGridView.AllowUserToAddRows = false;
                dataGridView.AutoSize = true;
                dataGridView.EditMode = DataGridViewEditMode.EditOnEnter;

                dataGridView.CellDoubleClick += dataGridView_CellDoubleClick;

                panel.Controls.Add(dataGridView);

                InitializeDataGridView();

                Button saveButton = new Button();
                saveButton.Text = "Сохранить";
                saveButton.Top = dataGridView.Bottom + 10;
                saveButton.Left = 20;
                saveButton.Click += SaveButton_Click;
                panel.Controls.Add(saveButton);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception: " + ex.Message);
            }
            finally
            {
                OleMessageFilter.Unregister();
            }
        }

        private void InitializeDataGridView()
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Value", typeof(string));
            dataTable.Columns.Add("Formula", typeof(string));
            dataTable.Columns.Add("Comment", typeof(string));

            foreach (Dimension dimension in dimensions)
            {
                dataTable.Rows.Add(dimension.DisplayName, dimension.Value.ToString(), dimension.Formula.ToString(), dimension.GetComment());
            }

            dataGridView.DataSource = dataTable;
            //dataGridView.Columns["Name"].ReadOnly = true;
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            SaveDimensionsToTable();
            InitializeDataGridView();
        }

        private void SaveDimensionsToTable()
        {
            DataTable dataTable = (DataTable)dataGridView.DataSource;

            foreach (DataRow row in dataTable.Rows)
            {
                string name = row["Name"].ToString();
                string value = row["Value"].ToString();
                string formula = row["Formula"].ToString();
                string comment = row["Comment"].ToString();

                Dimension dimension = FindDimensionByName(name);

                if (dimension != null)
                {
                    double doubleValue;
                    if (double.TryParse(value, out doubleValue))
                    {

                            dimension.SetComment(comment);
                            dimension.Value = doubleValue;
                            dimension.Formula = formula;

                    }
                    else
                    {
                        MessageBox.Show($"Неверный формат числа для измерения {name}");
                    }
                }
            }
        }

        private Dimension FindDimensionByName(string name)
        {
            if (dimensions != null)
            {
                foreach (Dimension dimension in dimensions)
                {
                    if (dimension.DisplayName == name)
                    {
                        return dimension;
                    }
                }
            }

            return null;
        }
        private void dataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView.Columns["Name"].Index && e.RowIndex >= 0)
            {
                string currentName = dataGridView.Rows[e.RowIndex].Cells["Name"].Value.ToString();

                string newName = ShowNameEditor(currentName);

                if (!string.IsNullOrEmpty(newName) && newName != currentName)
                {
                    UpdateDimensionName(currentName, newName);
                    InitializeDataGridView();
                }
            }
        }

        private string ShowNameEditor(string currentName)
        {
            Form nameEditorForm = new Form();
            nameEditorForm.Text = "Редактор имени";
            nameEditorForm.Size = new System.Drawing.Size(300, 150);
            nameEditorForm.StartPosition = FormStartPosition.CenterParent;

            Label label = new Label();
            label.Text = "Введите новое имя:";
            label.Location = new System.Drawing.Point(10, 20);

            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
            textBox.Text = currentName;
            textBox.Location = new System.Drawing.Point(10, 50);
            textBox.Size = new System.Drawing.Size(200, 20);

            Button okButton = new Button();
            okButton.Text = "OK";
            okButton.Location = new System.Drawing.Point(10, 80);
            okButton.Click += (sender, e) => nameEditorForm.Close();

            nameEditorForm.Controls.Add(label);
            nameEditorForm.Controls.Add(textBox);
            nameEditorForm.Controls.Add(okButton);

            nameEditorForm.ShowDialog();

            return textBox.Text;
        }

        private void UpdateDimensionName(string currentName, string newName)
        {
            Dimension dimension = FindDimensionByName(currentName);

            if (dimension != null)
            {
                dimension.VariableTableName = newName;
            }
        }

    }
}
