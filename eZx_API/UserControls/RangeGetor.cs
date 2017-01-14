using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;
using Point = System.Drawing.Point;
using TextBox = System.Windows.Forms.TextBox;

namespace eZx_API.UserControls
{
    /// <summary>
    /// 从 Excel 界面中获得指定的单元格区域
    /// </summary>
    public partial class RangeGetor : UserControl
    {
        private Application _excelApp;

        /// <summary> 为控件设置一个 Application 对象，此方法必须在构造函数执行后立即执行。 </summary>
        public void SetApplication(Application excelApp)
        {
            if (_excelApp == null) // 只能设置一次
            {
                if (excelApp != null)
                {
                    _excelApp = excelApp;
                }
                else
                {
                    throw new NullReferenceException("the excel application object can not be null.");
                }
            }
        }

        #region --- 控件属性

        private string _labeltext;

        [Category("RangeGetor"), Browsable(true), DefaultValue(""), Description("数据源的说明")]
        public string LabelText
        {
            set
            {
                label1.Text = value;
                _labeltext = value;
            }
            get { return _labeltext; }
        }

        private string _buttontext;

        [Category("RangeGetor"), Browsable(true), DefaultValue(""), Description("选择数据源的按钮上的简单字符")]
        public string ButtonText
        {
            set
            {
                button1.Text = value;
                _buttontext = value;
            }
            get { return _buttontext; }
        }

        #endregion

        /// <summary> 构造函数，在初始化对象后，必须要立即通过<seealso cref="SetApplication"/>方法赋值。 </summary>
        public RangeGetor()
        {
            InitializeComponent();
        }

        #region --- InitializeComponent

        private Button button1;
        private Label label1;
        private TextBox textBox1;

        private void InitializeComponent()
        {
            button1 = new Button();
            label1 = new Label();
            textBox1 = new TextBox();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            button1.Location = new Point(193, 0);
            button1.Margin = new Padding(2);
            button1.Name = "button1";
            button1.Size = new Size(17, 20);
            button1.TabIndex = 16;
            button1.UseVisualStyleBackColor = true;
            button1.Click += new EventHandler(button_srcX_Click);
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(1, 4);
            label1.Margin = new Padding(2, 0, 2, 0);
            label1.Name = "label1";
            label1.Size = new Size(29, 12);
            label1.TabIndex = 15;
            label1.Text = "说明";
            // 
            // textBox1
            // 
            textBox1.Anchor = ((AnchorStyles)((AnchorStyles.Left | AnchorStyles.Right)));
            textBox1.Location = new Point(56, 0);
            textBox1.Margin = new Padding(2);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(134, 21);
            textBox1.TabIndex = 14;
            textBox1.TextChanged += new EventHandler(textBox1_TextChanged);
            textBox1.Enter += new EventHandler(textBox_srcX_Enter);
            // 
            // RangeGetor
            // 
            AutoScaleDimensions = new SizeF(6F, 12F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(button1);
            Controls.Add(label1);
            Controls.Add(textBox1);
            Name = "RangeGetor";
            Size = new Size(212, 22);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private void textBox_srcX_Enter(object sender, EventArgs e)
        {
            if (Range != null)
            {
                Range.Select();
            }
        }

        private void button_srcX_Click(object sender, EventArgs e)
        {
            var inputResult = _excelApp.InputBox(
                Prompt: "选择单元格区域",
                Title: "选择",
                Default: (Range != null) ? Range.Address : "A1",
                Type: 8);

            if (inputResult is Range)
            {
                InnerSetRange(inputResult);
            }
            else // 如果选择单元格不成功，则不会返回 range 对象
            {
                return;
            }
        }

        #region --- textBox1_TextChanged 事件

        private bool _isInTextEditState;
        private string _originalText;

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Focused) // 说明是用户通过界面中敲击键盘来修改字符
            {
                _isInTextEditState = true;
                string newString = textBox1.Text;
                if (_originalText != newString)
                {
                    try
                    {
                        Worksheet sht = _excelApp.ActiveSheet;
                        Range rg = sht.Range[newString];

                        // 将新的 range 进行赋值
                        InnerSetRange(rg);
                        //
                        Range.Select();
                        //
                        _originalText = newString;
                    }
                    catch (Exception)
                    {
                        InnerSetRange(null);
                        return;
                    }
                    _isInTextEditState = false;
                }
            }
            else
            {
                // 说明是在代码中通过 textBox1.Text = "" 来进行整体赋值
            }
        }

        #endregion

        #region --- Range 对象的操作

        /// <summary>
        /// 当 <seealso cref="Range"/> 属性发生变化时触发
        /// </summary>
        [Category("RangeGetor"), Browsable(true), DefaultValue(""), Description("当控件所对应的Excel Range 范围发生变化时触发")]
        public event Action<object, Range> RangeChanged;

        private Range _range;

        /// <summary> 控件所对应的Excel单元格区域 </summary>
        [Browsable(false)]
        public Range Range
        {
            get { return _range; }
            private set
            {
                _range = value;
                ChangeTextWithoutRaiseEvent((value == null) ? "" : value.Address);
            }
        }

        /// <summary>
        /// 在类外部对 Range 对象的值进行设置
        /// </summary>
        /// <param name="newRange"></param>
        private void InnerSetRange(Range newRange)
        {
            _rangeHasBeenChangeByOuterEvent = false;

            RaisePossibleEvent(newRange);

            if (!_rangeHasBeenChangeByOuterEvent)
            {
                // Range 属性的赋值一定要在 触发事件之后，因为赋值之后 新旧Range就一样了。
                Range = newRange;
            }
        }

        /// <summary>
        /// 如果 Range 属性已经通过<seealso cref="SetRange"/>方法被外部用户强制修改过了，
        /// 那么在<seealso cref="InnerSetRange"/>方法中就不能再将其值复原了。
        /// </summary>
        private bool _rangeHasBeenChangeByOuterEvent = false;

        /// <summary>
        /// 在类外部对 Range 对象的值进行设置
        /// </summary>
        /// <param name="newRange"></param>
        /// <param name="isOuterEvent"> 此方法是否通过外部的 <seealso cref="RangeChanged"/>事件触发时被执行 </param>
        /// <param name="raisePossibleEvent"> 是否触发可能的 <seealso cref="RangeChanged"/> 事件 </param>
        public void SetRange(Range newRange, bool isOuterEvent, bool raisePossibleEvent = true)
        {
            if (raisePossibleEvent)
            {
                RaisePossibleEvent(newRange);
            }

            // Range 属性的赋值一定要在 触发事件之后，因为赋值之后 新旧Range就一样了。
            Range = newRange;
            //
            _rangeHasBeenChangeByOuterEvent = isOuterEvent;
        }

        private void RaisePossibleEvent(Range newRange)
        {
            bool valueChanged = (newRange == null && _range != null) ||
                                (newRange != null && _range == null) ||
                                (newRange != null) && (_range != null) && (newRange.Address != _range.Address);

            if (valueChanged) // 如果前后 Range 范围相同，则相当于未进行设置
            {
                if (RangeChanged != null) RangeChanged(this, newRange);
            }
        }

        private void ChangeTextWithoutRaiseEvent(string newString)
        {
            if (!_isInTextEditState)
            {
                textBox1.TextChanged -= textBox1_TextChanged;

                textBox1.Text = newString;

                textBox1.TextChanged += textBox1_TextChanged;
            }
        }

        #endregion
    }
}