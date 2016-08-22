using System.Collections.Generic;
using System;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.VisualBasic;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Linq;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using eZexcelAPI;

namespace eZx.Database
{

    public partial class Form_ConstructDatabase
    {

        private BindingList<DataField> List_FieldInfo = new BindingList<DataField>();
        public Excel.Worksheet Worksheet { get; set; }
        /// <summary>
        /// 当前窗口是否处于“构造数据库”的模式，如果为False，则为“编辑数据库”的模式
        /// </summary>
        private bool IsConstructingMode;

        /// <summary>
        /// Worksheet.UsedRange.Value所返回的值，此二维数组中，左上角的第一个元素的下标值为(1,1)
        /// </summary>
        /// <remarks>此二维数组中包含了字段信息以及每一个字段中的数据</remarks>
        private object[,] F_DataValue;

        ///<summary> 此字段名称本身的数据类型。
        /// 一般情况下，一个字段的名称只要是一个字符就可以了，但是它也可以代表具体的含义，比如在具体某一天的日期“2016/2/6” </summary>
        private DataField.eZDataType F_FieldType;

        eZDataSheet DataSheet = null;

        #region   ---  构造函数与窗体的加载、打开与关闭

        /// <summary> 构造函数 </summary>
        /// <param name="Sheet"></param>
        /// <param name="ConstructingMode">当前窗口是否处于“构造数据库”的模式，如果为False，则为“编辑数据库”的模式</param>
        ///<param name="DataSheet">当以“构造数据库”模式打开时，此参数可不赋值；当以“编辑数据库”模式打开式，此参数为对应的活动数据库。</param>
        /// <remarks></remarks>
        public Form_ConstructDatabase(Excel.Worksheet Sheet, bool ConstructingMode, eZDataSheet DataSheet = null)
        {
            // This call is required by the designer.
            InitializeComponent();
            // Add any initialization after the InitializeComponent() call.
            List_FieldInfo.AllowNew = true; // .Add(New DataField("", eZDataType.字符, False, eZDataType.字符))
                                            //
            SetupDataGridView();
            //
            this.ComboBox_CommonDataType.DataSource = Enum.GetValues(typeof(DataField.eZDataType));
            this.ComboBox_FieldType.DataSource = Enum.GetValues(typeof(DataField.eZDataType));
            //
            this.Worksheet = Sheet;
            this.IsConstructingMode = ConstructingMode;
            this.DataSheet = DataSheet;
        }

        private void SetupDataGridView()
        {
            this.eZDataGridView1.AutoGenerateColumns = false;
            this.eZDataGridView1.AllowUserToAddRows = true;
            this.eZDataGridView1.AutoSize = true;
            //
            // 添加数据列
            DataGridViewTextBoxColumn Column_FieldName = new DataGridViewTextBoxColumn();
            Column_FieldName.DataPropertyName = "Name";
            Column_FieldName.HeaderText = "名称";
            Column_FieldName.Resizable = DataGridViewTriState.False;
            this.eZDataGridView1.Columns.Add(Column_FieldName);
            //
            DataGridViewComboBoxColumn Column_DataType = new DataGridViewComboBoxColumn();
            Column_DataType.DataSource = Enum.GetValues(typeof(DataField.eZDataType)); //对于ComboBoxColumn，这一句是必须的。
            Column_DataType.DataPropertyName = "DataType";
            Column_DataType.Name = "DataType";
            Column_DataType.HeaderText = "数据类型";
            Column_DataType.Width = 70;
            Column_DataType.Resizable = DataGridViewTriState.False;
            this.eZDataGridView1.Columns.Add(Column_DataType);
            //
            DataGridViewCheckBoxColumn Column_NullAllowed = new DataGridViewCheckBoxColumn();
            Column_NullAllowed.DataPropertyName = "NullAllowed";
            Column_NullAllowed.Width = 70;
            Column_NullAllowed.HeaderText = "允许空值";
            Column_NullAllowed.Resizable = DataGridViewTriState.False;
            this.eZDataGridView1.Columns.Add(Column_NullAllowed);
            //
            DataGridViewButtonColumn Column_Check = new DataGridViewButtonColumn();
            Column_Check.HeaderText = "检验";
            Column_Check.Name = "CheckField";
            Column_Check.Text = "Check Field";
            // Use the Text property for the button text for all cells rather
            // than using each cell's value as the text for its own button.
            Column_Check.UseColumnTextForButtonValue = true;
            Column_Check.Resizable = DataGridViewTriState.False;
            this.eZDataGridView1.Columns.Insert(0, Column_Check);
        }

        public eZDataSheet ShowDialog()
        {
            DialogResult res = base.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.Yes)
            {
                // 构造数据库并返回
                DataSheet = new eZDataSheet(Worksheet, List_FieldInfo, this.F_FieldType);

                return DataSheet;
            }
            else
            {
                return null;
            }
        }

        // 加载窗口: 每次在Form.ShowDialog方法中，均会触发此Load事件
        public void Form_ConstructDatabase_Load(object sender, EventArgs e)
        {
            if (IsConstructingMode)
            {
                ConstructDataBase();
            }
            else
            {
                EditDataBase(DataSheet);
            }
        }

        // 关闭窗口
        public void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion

        /// <summary>
        /// 构造数据库
        /// </summary>
        /// <remarks></remarks>
        private void ConstructDataBase()
        {
            Range rg = Worksheet.UsedRange;
            rg.Select();
            if (rg.Cells[1, 1].Address != Worksheet.Cells[1, 1].Address)
            {
                MessageBox.Show("数据表的第一行/列没有数据", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //
            this.List_FieldInfo.Clear();
            if (rg.Cells.Count > 1)
            {
                this.F_DataValue = rg.Value;
                UInt16 FieldsCount =(UInt16) F_DataValue.GetUpperBound(1);// (UInt16)Information.UBound((System.Array)F_DataValue, 2);
                //
                //
                string FieldName = "";
                for (UInt16 FieldIndex = 1; FieldIndex <= FieldsCount; FieldIndex++)
                {
                    FieldName = System.Convert.ToString(F_DataValue[1, FieldIndex]);
                    List_FieldInfo.Add(new DataField(FieldName, FieldIndex));
                }
            }
            else // 说明只选择了一个单元格，此时rg.Value并不会返回一个数组，而是返回一个String或Double等的值
            {
                List_FieldInfo.Add(new DataField(rg.Value.ToString(), 1));
            }
            this.eZDataGridView1.DataSource = List_FieldInfo;
        }

        /// <summary>
        /// 编辑数据库
        /// </summary>
        /// <remarks></remarks>
        private void EditDataBase(eZDataSheet DataSheet)
        {
            this.Text = "编辑数据库";
            // 每个字段的数据类型
            this.eZDataGridView1.DataSource = DataSheet.Fields;

            // 字段名称的数据类型

        }

        #region   ---  检验字段的信息
        /// <summary>
        /// 同时检验一个字段的名称的数据类型，以及此字段的此列数据的数据类型
        /// </summary>
        /// <param name="Field">某一个字段</param>
        /// <param name="Value">整个数据表的数据（包含字段），数组中的第一个元素的下标为1</param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool ValidateField(DataField Field, ref object[,] Value)
        {
            bool blnIsValidated = true;
            if (!ValidateFieldType(Field))
            {
                return false;
            }
            if (!ValidateFieldDataType(Field, Value))
            {
                return false;
            }
            return blnIsValidated;
        }

        /// <summary>
        /// 检验某一字段的一列数据的数据类型
        /// </summary>
        /// <param name="Field">字段信息</param>
        /// <param name="Value">整个数据表的数据（包含字段），数组中的第一个元素的下标为1</param>
        /// <returns></returns>
        private bool ValidateFieldDataType(DataField Field, object[,] Value)
        {
            bool blnIsValidated = true;

            UInt32 DataCount = (UInt32)(Value.GetUpperBound(0) - 1); // 数据的个数（不包括字段名称）
            object v = null;
            if (Field.NullAllowed) // 允许空值
            {
                for (UInt32 i = 2; i <= DataCount; i++)
                {
                    v = Value[i, Field.ColumnIndex];
                    if ((v != null) && (!DataField.IsCompatible(System.Convert.ToString(v), Field.DataType)))
                    {
                        return false;
                    }
                }
            }
            else // 不允许空值
            {
                for (UInt32 i = 2; i <= DataCount; i++)
                {
                    if (!DataField.IsCompatible(System.Convert.ToString(Value[i, Field.ColumnIndex]), Field.DataType))
                    {
                        return false;
                    }
                }

            }


            return blnIsValidated;
        }

        /// <summary>
        /// 检查某一字段的名称本身的数据类型
        /// </summary>
        /// <param name="Field"></param>
        /// <returns></returns>
        private bool ValidateFieldType(DataField Field)
        {
            DataField with_1 = Field;
            if (DataField.IsCompatible(with_1.Name, F_FieldType))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region    ---  事件处理

        /// <summary> 点击表格控件中的单元格中的对象 </summary>
        public void eZDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0) // 说明点击的是“检验字段”的按钮
            {
                DataField FieldDt = (DataField)(eZDataGridView1.Rows[e.RowIndex].DataBoundItem);
                if (FieldDt.ColumnIndex == 1) // 第一个字段只检验数据的类型，而不检查字段名称本身的类型
                {
                    if (ValidateFieldDataType(FieldDt, this.F_DataValue))
                    {
                        MessageBox.Show("字段检验合格", "Congratulations", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    else
                    {
                        MessageBox.Show("字段检验不合格", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    if (ValidateField(FieldDt, ref this.F_DataValue))
                    {
                        MessageBox.Show("字段检验合格", "Congratulations", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    else
                    {
                        MessageBox.Show("字段检验不合格", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// ! 检验所有的字段，完成数据库的构造或者编辑
        /// </summary>
        public void CheckAllFields(object sender, EventArgs e)
        {
            DataField FieldDt = default(DataField);
            bool blnOk = true;
            if (List_FieldInfo.Count > 0)
            {

                // 从第二个字段开始来检验字段名称的数据类型，因为对于“字段名称本身的数据类型”的检验，是不包括第一个字段的。
                for (int Index = 1; Index <= List_FieldInfo.Count - 1; Index++)
                {
                    FieldDt = this.List_FieldInfo[Index];
                    if (!ValidateFieldType(FieldDt))
                    {
                        MessageBox.Show(string.Format("第{0}个字段： {1} 的字段名称的数据类型检验不合格", Index + 1, FieldDt.Name),
                            "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        blnOk = false;
                        break;
                    }
                }

                // 从第一个字段开始来检验每一列数据的数据类型
                if (blnOk)
                {
                    for (int Index = 0; Index <= List_FieldInfo.Count - 1; Index++)
                    {
                        FieldDt = this.List_FieldInfo[Index];
                        if (!ValidateFieldDataType(FieldDt, this.F_DataValue))
                        {
                            MessageBox.Show(string.Format("第{0}个字段： {1} 的数据检验不合格", Index + 1, FieldDt.Name),
                                "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            blnOk = false;
                            break;
                        }
                    }

                }
                //
            }
            if (blnOk)
            {
                this.DialogResult = System.Windows.Forms.DialogResult.Yes;
                MessageBox.Show("所有字段检验合格",
                    "Congratulations", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                this.Close();
            }
            else
            {
                //   Me.DialogResult = System.Windows.Forms.DialogResult.No
            }
        }

        // 错误处理
        public void eZDataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            var aa = this.eZDataGridView1[e.ColumnIndex, e.RowIndex].ValueType;
            var a = this.eZDataGridView1[e.ColumnIndex, e.RowIndex].Value;

            MessageBox.Show(e.Exception.Message + "\r\n" +
                "行号：" + System.Convert.ToString(e.RowIndex) + "\r\n" +
                "列号：" + System.Convert.ToString(e.ColumnIndex) + "\r\n" +
                e.Context.ToString());
            e.Cancel = true;
        }
        // 改变基本数据类型
        public void ChangeAllFieldDataType(object sender, EventArgs e)
        {
            DataField.eZDataType ty = (DataField.eZDataType)this.ComboBox_CommonDataType.SelectedValue;
            UInt32 Count = (UInt16)this.eZDataGridView1.Rows.Count;
            for (int r = 0; r <= Count - 1; r++)
            {
                this.eZDataGridView1["DataType", r].Value = ty;
            }
        }
        // 改变字段名称本身的数据类型
        public void ChangeFieldType(object sender, EventArgs e)
        {
            //  Dim blnSucceed As Boolean
            DataField.eZDataType ezTp = (DataField.eZDataType)this.ComboBox_FieldType.SelectedValue;
            this.F_FieldType = ezTp;
            // 更新界面
            if (ezTp == DataField.eZDataType.字符)
            {
                this.CheckBox1.CheckState = CheckState.Indeterminate;
                this.CheckBox1.Enabled = false;
            }
            else
            {
                this.CheckBox1.CheckState = CheckState.Checked;
                this.CheckBox1.Enabled = true;
            }

            // 检验字段
            string FieldName = "";
            for (UInt32 FieldIndex = 1; FieldIndex <= List_FieldInfo.Count - 1; FieldIndex++) // 不检验第一个字段的数据类型
            {
                DataField df = List_FieldInfo[(UInt16)FieldIndex];
                FieldName = df.Name;
                if (!ValidateFieldType(df))
                {
                    MessageBox.Show("第" + System.Convert.ToString(df.ColumnIndex) + "个字段名称不符合指定的数据类型：" + FieldName,
                        "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    //选择出错的那一行
                    this.eZDataGridView1.Rows[df.ColumnIndex - 1].Selected = true;
                    break;
                }
            }
        }

        // 改变每个字段“是否允许空值”
        public void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            bool blnAllowNull = CheckBox2.Checked;
            //
            foreach (DataField df in this.List_FieldInfo)
            {
                df.NullAllowed = blnAllowNull;
            }
            // 刷新界面显示
            this.eZDataGridView1.Refresh();
        }
        // 添加新的数据行
        public void FieldInfo_AddingNew(object sender, AddingNewEventArgs e)
        {
            e.NewObject = new DataField("字段名称", (ushort)DataField.eZDataType.字符, DataField.eZDataType.字符, true);
        }

        #endregion

    }
}
