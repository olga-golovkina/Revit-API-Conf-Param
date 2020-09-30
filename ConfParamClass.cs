using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Form = System.Windows.Forms.Form;
using TextBox = System.Windows.Forms.TextBox;
using View = System.Windows.Forms.View;

namespace ConfParam
{
    [Transaction(TransactionMode.Manual)]

    class ConfParamClass : IExternalCommand
    {
        Form formCategories, formParam, textDialog;
        TextBox textBox;
        TableLayoutPanel tLPCatMain, tLPParamMain;
        FlowLayoutPanel fLPCatBottom, fLPParamBottom, fLParamRightPanelButtons;
        Label lHeaderCat, lHeaderParamLeft, lHeaderParamRight, lConfParam;
        ListView lVCategories, lVParamLeft;
        Button bCatCancel, bCatNext, bParamOK, bParamBack, bParamCancel, bAddEdParam, bAddParam, bDelParam, bClearParam, bSetWordParam;
        List<string> checkedCat = new List<string>();
        List<Category> checkedCategory = new List<Category>();
        List<Element> existType = new List<Element>();
        int countPar, gEx, gType;
        bool newParam, typeParam;
        string word = ", ";
        List<List<ParameterSet>> allSelectedParamsExempl = new List<List<ParameterSet>>();
        List<List<ParameterSet>> allSelectedParamsType = new List<List<ParameterSet>>();

        List<ParameterSet> allParamExempl = new List<ParameterSet>();
        List<ParameterSet> allParamType = new List<ParameterSet>();

        List<Parameter> generalParamsExempl = new List<Parameter>();
        List<Parameter> generalParamsType = new List<Parameter>();

        List<int> fParamIdExempl = new List<int>();
        List<int> fParamIdType = new List<int>();
        List<int> generParamIdExempl = new List<int>();
        List<int> generParamIdType = new List<int>();


        Parameter paramEd;
        List<Parameter> lstParametersForConfig = new List<Parameter>();
       // IList<(Parameter, List<Parameter>, List<Element>)> lst = new IList<(Parameter, List<Parameter>, List<Element>)>();

        UIDocument UIDoc;
        Document document;

        //Форма выбора категории
        #region
        private void createFormCategories() 
        {
            formCategories = new Form();
            setFormCategories();
            formCategories.AcceptButton = bCatNext;
            formCategories.Show();
        }

        private void setFormCategories()
        {
            formCategories.StartPosition = FormStartPosition.CenterScreen;
            formCategories.Width = 500;
            formCategories.Height = 500;
            formCategories.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            formCategories.Text = "Configuration Parameters";
            setLayoutFormCategories();
        }

        private void setLayoutFormCategories()
        {
            formCategories.SuspendLayout();

            createTLPCatMain();
            formCategories.Controls.Add(tLPCatMain);

            formCategories.ResumeLayout();
            formCategories.Refresh();
        }

        private void createTLPCatMain() 
        {
            tLPCatMain = new TableLayoutPanel();
            tLPCatMain.RowCount = 3;
            tLPCatMain.Dock = DockStyle.Fill;

            setStyleTLPCatMain();
            setLayoutTLPCatMain();
        }

        private void setStyleTLPCatMain()
        {
            RowStyle rowStyle01, rowStyle02, rowStyle03;
            
            rowStyle01 = new RowStyle();
            rowStyle02 = new RowStyle();
            rowStyle03 = new RowStyle();

            rowStyle01.SizeType = SizeType.Percent;
            rowStyle02.SizeType = SizeType.Percent;
            rowStyle03.SizeType = SizeType.Percent;

            rowStyle01.Height = 9;
            rowStyle02.Height = 81;
            rowStyle03.Height = 10;

            tLPCatMain.RowStyles.Add(rowStyle01);
            tLPCatMain.RowStyles.Add(rowStyle02);
            tLPCatMain.RowStyles.Add(rowStyle03);
        }

        private void setLayoutTLPCatMain()
        {
            tLPCatMain.SuspendLayout();
            
            createLHeaderCat();
            createLVCategories();
            createFLPCatBottom();

            tLPCatMain.Controls.Add(lHeaderCat, 0, 0);
            tLPCatMain.Controls.Add(lVCategories, 0, 1);
            tLPCatMain.Controls.Add(fLPCatBottom, 0, 2);

            tLPCatMain.ResumeLayout();
            tLPCatMain.Refresh();
        }

        private void createLHeaderCat()
        {
            lHeaderCat = new Label();
            lHeaderCat.Text = "Select Category";
            lHeaderCat.Dock = DockStyle.Fill;
            lHeaderCat.Margin = new Padding(7, 10, 0, 0);
        }

        private void createLVCategories()
        {
            lVCategories = new ListView();
            lVCategories.CheckBoxes = true;
            lVCategories.Sorting = SortOrder.Ascending;
            lVCategories.View = View.Details;
            lVCategories.HeaderStyle = ColumnHeaderStyle.None;
            lVCategories.MultiSelect = true;
            lVCategories.Dock = DockStyle.Fill;
            lVCategories.Scrollable = true;
            lVCategories.Margin = new Padding(10, 0, 10, 0);
            lVCategories.ItemCheck += lVCategories_ItemCheck;

            fillingLVCategories();
        }

        private void lVCategories_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (lVCategories.CheckedItems.Count == 0)
                bCatNext.Enabled = false;
            else
                bCatNext.Enabled = true;
        }

        private void fillingLVCategories()
        {
            lVCategories.Columns.Add("Category");
            lVCategories.Columns[0].Width = -2;
            foreach (Category elem in getAllCategories())
                if (elem.CategoryType == CategoryType.Model || elem.CategoryType == CategoryType.AnalyticalModel)
                    lVCategories.Items.Add(elem.Name);
            checkingLVCategoriesItems();
        }

        private void checkingLVCategoriesItems()
        {
            foreach (string elem in checkedCat)
                foreach (ListViewItem item in lVCategories.Items)
                    if (elem == item.Text)
                        item.Checked = true;
        }

        private void createFLPCatBottom()
        {
            fLPCatBottom = new FlowLayoutPanel();
            fLPCatBottom.FlowDirection = FlowDirection.RightToLeft;
            fLPCatBottom.Dock = DockStyle.Fill;
            setLayoutFLPCatBottom();
        }

        private void setLayoutFLPCatBottom()
        {
            createButtonsCat();

            fLPCatBottom.SuspendLayout();

            fLPCatBottom.Controls.Add(bCatNext);
            fLPCatBottom.Controls.Add(bCatCancel);

            fLPCatBottom.ResumeLayout();
            fLPCatBottom.Refresh();
        }
        private void createButtonsCat()
        {
            bCatNext = new Button();
            bCatCancel = new Button();

            bCatNext.Text = "Next >";
            bCatCancel.Text = "Cancel";

            bCatNext.Anchor = AnchorStyles.Right;
            bCatCancel.Anchor = AnchorStyles.Right;

            bCatNext.Margin = new Padding(0, 10, 5, 0);
            bCatCancel.Margin = new Padding(0, 10, 5, 0);

            bCatNext.Click += bCatNext_Click;
            bCatCancel.Click += bCatCancel_Click;

            bCatNext.Enabled = false;
        }

        private void bCatNext_Click(object sender, EventArgs e)
        {
            var checkedILV = lVCategories.CheckedItems;

            checkedCat.Clear();

            foreach (ListViewItem elem in checkedILV)
                checkedCat.Add(elem.Text);

            foreach (string str in checkedCat)
                foreach (Category elem in getAllCategories())
                    if (elem.Name == str)
                        checkedCategory.Add(elem);
                    
            int id = -1002050;

            foreach (Category category in checkedCategory)
            {
                foreach (Element elem in getFamilies(category))
                {
                    allParamExempl.Add(elem.Parameters);
                    allSelectedParamsExempl.Add(allParamExempl);
                    allExempl.Add(elem);
                    
                    var paramType = elem.get_Parameter((BuiltInParameter)id);

                    foreach (Element element in getTypeFamilies(category))
                        if (paramType.AsValueString() == element.Name)
                        {
                            allParamType.Add(element.Parameters);
                            allSelectedParamsType.Add(allParamType);
                            allType.Add(element);

                            //Trace.WriteLine(element.Name);
                            existType.Add(element);
                        }
                }
            }

            collectGenParametersExempl();
            getGeneralParamExempl();

            collectGenParametersType();
            getGeneralParamType();

            generParamIdExempl.Clear();
            generParamIdType.Clear();

            generalParamsExempl = generalParamsExempl.GroupBy(param => param.Id).Select(param => param.First()).ToList();
            generalParamsType = generalParamsType.GroupBy(param => param.Id).Select(param => param.First()).ToList();

            formCategories.Close();
            createFormParam();
        }

        private void bCatCancel_Click(object sender, EventArgs e)
        {
            formCategories.Close();
        }

        #endregion //Форма выбора категории

        //Форма конфигуратора параметров
        #region
        private void createFormParam()
        {
            formParam = new Form();
            setFormParam();
            formParam.AcceptButton = bParamOK;
            formParam.Show();
        }

        private void setFormParam()
        {
            formParam.StartPosition = FormStartPosition.CenterScreen;
            formParam.Width = 700;
            formParam.Height = 500;
            formParam.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            formParam.Text = "Configuration Parameters";
            setLayoutFormParam();
        }

        private void setLayoutFormParam()
        {
            formParam.SuspendLayout();

            createTLPParamMain();
            formParam.Controls.Add(tLPParamMain);

            formParam.ResumeLayout();
            formParam.Refresh();
        }

        private void createTLPParamMain()
        {
            tLPParamMain = new TableLayoutPanel();
            tLPParamMain.RowCount = 3;
            tLPParamMain.ColumnCount = 3;
            tLPParamMain.Dock = DockStyle.Fill;

            setStyleTLPParamMain();
            setLayoutTLPParamMain();
        }

        private void setStyleTLPParamMain()
        {
            RowStyle rowStyle01, rowStyle02, rowStyle03;
            ColumnStyle columnStyle01, columnStyle02, columnStyle03;

            rowStyle01 = new RowStyle();
            rowStyle02 = new RowStyle();
            rowStyle03 = new RowStyle();
            columnStyle01 = new ColumnStyle();
            columnStyle02 = new ColumnStyle();
            columnStyle03 = new ColumnStyle();

            rowStyle01.SizeType = SizeType.Percent;
            rowStyle02.SizeType = SizeType.Percent;
            rowStyle03.SizeType = SizeType.Percent;
            columnStyle01.SizeType = SizeType.Percent;
            columnStyle02.SizeType = SizeType.Percent;
            columnStyle03.SizeType = SizeType.Percent;

            rowStyle01.Height = 9;
            rowStyle02.Height = 81;
            rowStyle03.Height = 10;
            columnStyle01.Width = 45;
            columnStyle02.Width = 45;
            columnStyle03.Width = 10;

            tLPParamMain.RowStyles.Add(rowStyle01);
            tLPParamMain.RowStyles.Add(rowStyle02);
            tLPParamMain.RowStyles.Add(rowStyle03);
            tLPParamMain.ColumnStyles.Add(columnStyle01);
            tLPParamMain.ColumnStyles.Add(columnStyle02);
            tLPParamMain.ColumnStyles.Add(columnStyle03);
        }

        private void setLayoutTLPParamMain()
        {
            createLHeaderParamLeft();
            createLHeaderParamRight();
            createFLPParamBottom();
            createFLParamRightPanelButtons();
            createLVParamLeft();
            createLConfParam();

            tLPParamMain.SuspendLayout();

            tLPParamMain.Controls.Add(lHeaderParamLeft, 0, 0);
            tLPParamMain.Controls.Add(lHeaderParamRight, 1, 0);
            tLPParamMain.Controls.Add(lVParamLeft, 0, 1);
            tLPParamMain.Controls.Add(lConfParam, 1, 1);
            tLPParamMain.Controls.Add(fLPParamBottom, 1, 2);
            tLPParamMain.Controls.Add(fLParamRightPanelButtons, 2, 1);

            tLPParamMain.ResumeLayout();
            tLPParamMain.Refresh();
        }

        private void createLVParamLeft()
        {
            lVParamLeft = new ListView();
            lVParamLeft.Sorting = SortOrder.Ascending;
            lVParamLeft.View = View.Details;
            lVParamLeft.HeaderStyle = ColumnHeaderStyle.None;
            lVParamLeft.MultiSelect = false;
            lVParamLeft.Dock = DockStyle.Fill;
            lVParamLeft.Scrollable = true;
            lVParamLeft.Margin = new Padding(10, 0, 10, 0);

            newParam = false;

            fillingLVParamLeft();
        }

        private void fillingLVParamLeft()
        {

            List<string> textItems = new List<string>();
            try
            {
                lVParamLeft.Items.Clear();
            }
            catch (Exception) { }

            if (newParam == false)
            {
                lVParamLeft.Clear();

                lVParamLeft.Columns.Add("Param");
                lVParamLeft.Columns[0].Width = -2;

                foreach (Parameter elem in generalParamsExempl)
                    if (!elem.IsReadOnly && elem.StorageType != StorageType.None)
                        textItems.Add(elem.Definition.Name);

                foreach (Parameter elem in generalParamsType)
                    if (!elem.IsReadOnly && elem.StorageType != StorageType.None)
                        textItems.Add(elem.Definition.Name);
            }
            else
            {
                if (allExempl.First().get_Parameter((BuiltInParameter)paramEd.Id.IntegerValue) != null)
                {
                    typeParam = false;

                    if (paramEd.StorageType == StorageType.String)
                    {
                        foreach (Parameter elem in generalParamsExempl.Where(x => x.Definition.Name != paramEd.Definition.Name))
                            if (elem.HasValue)
                                textItems.Add(elem.Definition.Name);
                        foreach (Parameter elem in generalParamsType.Where(x => x.Definition.Name != paramEd.Definition.Name))
                            if (elem.HasValue)
                                textItems.Add(elem.Definition.Name);
                    }
                    else
                    {
                        foreach (Parameter elem in generalParamsExempl.Where(x => x.Definition.Name != paramEd.Definition.Name))
                            if (elem.IsReadOnly && elem.StorageType == paramEd.StorageType && elem.HasValue)
                                textItems.Add(elem.Definition.Name);
                        foreach (Parameter elem in generalParamsType.Where(x => x.Definition.Name != paramEd.Definition.Name))
                            if (elem.IsReadOnly && elem.StorageType == paramEd.StorageType && elem.HasValue)
                                textItems.Add(elem.Definition.Name);
                    }
                }

                else if (allType.First().get_Parameter((BuiltInParameter)paramEd.Id.IntegerValue) != null)
                {
                    typeParam = true;

                    if (paramEd.StorageType == StorageType.String)
                    {
                        foreach (Parameter elem in generalParamsType.Where(x => x.Definition.Name != paramEd.Definition.Name))
                            if (elem.HasValue)
                            {
                                textItems.Add(elem.Definition.Name);
                            }
                    }
                    else
                    {
                        foreach (Parameter elem in generalParamsType.Where(x => x.Definition.Name != paramEd.Definition.Name))
                            if (elem.IsReadOnly && elem.StorageType == paramEd.StorageType && elem.HasValue)
                            {
                                textItems.Add(elem.Definition.Name);
                            }
                    }
                }
                else return;
            }

            textItems = textItems.GroupBy(param => param).Select(param => param.First()).ToList();

            foreach (string itemText in textItems)
                lVParamLeft.Items.Add(itemText);

        }

        private void createLConfParam()
        {
            lConfParam = new Label();
            lConfParam.Dock = DockStyle.Top;
            lConfParam.Height = 20;
            lConfParam.Margin = new Padding(10, 0, 10, 0);
            lConfParam.BackColor = System.Drawing.Color.White;
            lConfParam.BorderStyle = BorderStyle.FixedSingle;
            lConfParam.AutoSize = true;
        }

        private void createLHeaderParamLeft()
        {
            lHeaderParamLeft = new Label();
            lHeaderParamLeft.Text = "Available parameters for editing";
            lHeaderParamLeft.Dock = DockStyle.Fill;
            lHeaderParamLeft.Margin = new Padding(7, 10, 0, 0);
        }

        private void createLHeaderParamRight()
        {
            lHeaderParamRight = new Label();
            lHeaderParamRight.Text = "Parameter configurator";
            lHeaderParamRight.Dock = DockStyle.Fill;
            lHeaderParamRight.Margin = new Padding(7, 10, 0, 0);
        }

        private void createFLPParamBottom()
        {
            fLPParamBottom = new FlowLayoutPanel();
            fLPParamBottom.FlowDirection = FlowDirection.RightToLeft;
            fLPParamBottom.Dock = DockStyle.Fill;
            setLayoutFLPParamBottom();
        }

        private void setLayoutFLPParamBottom()
        {
            createButtonsParam();

            fLPParamBottom.SuspendLayout();

            fLPParamBottom.Controls.Add(bParamOK);
            fLPParamBottom.Controls.Add(bParamBack);
            fLPParamBottom.Controls.Add(bParamCancel);

            fLPParamBottom.ResumeLayout();
            fLPParamBottom.Refresh();
        }

        private void createButtonsParam()
        {
            
            //Нижние кнопки
            #region
            bParamOK = new Button();
            bParamBack = new Button();
            bParamCancel = new Button();

            bParamOK.Text = "OK";
            bParamBack.Text = "< Back";
            bParamCancel.Text = "Cancel";

            bParamOK.Anchor = AnchorStyles.Right;
            bParamBack.Anchor = AnchorStyles.Right;
            bParamCancel.Anchor = AnchorStyles.Right;

            bParamOK.Margin = new Padding(0, 10, 5, 0);
            bParamBack.Margin = new Padding(0, 10, 5, 0);
            bParamCancel.Margin = new Padding(0, 10, 5, 0);

            bParamOK.Click += bParamOK_Click;
            bParamBack.Click += bParamBack_Click;
            bParamCancel.Click += bParamCancel_Click;

            bParamOK.Enabled = false;
            #endregion

            //Кнопки для правой панели
            #region
            bAddEdParam = new Button();
            bAddParam = new Button();
            bDelParam = new Button();
            bClearParam = new Button();
            bSetWordParam = new Button();

            bAddEdParam.Text = "AEP";
            bAddParam.Text = "AP";
            bDelParam.Text = "DP";
            bClearParam.Text = "CP";
            bSetWordParam.Text = "SW";

            bAddEdParam.Margin = new Padding(5, 75, 0, 0);
            bAddParam.Margin = new Padding(5, 10, 0, 0);
            bDelParam.Margin = new Padding(5, 10, 0, 0);
            bClearParam.Margin = new Padding(5, 10, 0, 0);
            bSetWordParam.Margin = new Padding(5, 10, 0, 0);

            bAddEdParam.Height = 25;
            bAddParam.Height = 25;
            bDelParam.Height = 25;
            bClearParam.Height = 25;
            bSetWordParam.Height = 25;

            bAddEdParam.Width = 35;
            bAddParam.Width = 35;
            bDelParam.Width = 35;
            bClearParam.Width = 35;
            bSetWordParam.Width = 35;

            bAddEdParam.Click += bAddEdParam_Click;
            
            bAddParam.Click += bAddParam_Click;
            bDelParam.Click += bDelParam_Click;
            bClearParam.Click += bClearParam_Click;
            bSetWordParam.Click += bSetWordParam_Click;

            bAddParam.Enabled = false;
            bDelParam.Enabled = false;
            bClearParam.Enabled = false;

            ToolTip ToolTip1 = new ToolTip();
            ToolTip1.SetToolTip(bAddEdParam, "Кнопка для выбора параметра, над которым будет проведена конфигурация");
            ToolTip1.SetToolTip(bAddParam, "Кнопка для выбора параметра, который будет использоваться в конфигурации");
            ToolTip1.SetToolTip(bDelParam, "Кнопка для удаление последнего параметра, который используется в конфигурации");
            ToolTip1.SetToolTip(bClearParam, "Очистка конфигуратора");
            ToolTip1.SetToolTip(bSetWordParam, "Задание разделителя. По умолчанию разделить будет , ");

            #endregion  
        }

        Button bOk, bCancel;
        private void bSetWordParam_Click(object sender, EventArgs e)
        {
            textDialog = new Form();

            textDialog.StartPosition = FormStartPosition.CenterParent;
            textDialog.Width = 250;
            textDialog.Height = 145;

            textDialog.FormBorderStyle = FormBorderStyle.FixedToolWindow;

            TableLayoutPanel tableLayoutPanel = new TableLayoutPanel();
            
            tableLayoutPanel.RowCount = 3;
            tableLayoutPanel.ColumnCount = 1;
            tableLayoutPanel.Dock = DockStyle.Fill;

            tableLayoutPanel.Padding = new Padding(0);
            tableLayoutPanel.Margin = new Padding(0);

            RowStyle row01, row02, row03;
            row01 = new RowStyle();
            row02 = new RowStyle();
            row03 = new RowStyle();

            row01.SizeType = SizeType.Percent;
            row02.SizeType = SizeType.Percent;
            row03.SizeType = SizeType.Percent;

            row01.Height = 33;
            row02.Height = 33;
            row03.Height = 33;

            tableLayoutPanel.RowStyles.Add(row01);
            tableLayoutPanel.RowStyles.Add(row02);
            tableLayoutPanel.RowStyles.Add(row03);

            textDialog.Controls.Add(tableLayoutPanel);

            Label l01 = new Label();
            l01.Text = "Enter separator: ";
            l01.Anchor = AnchorStyles.Top;
            l01.Dock = DockStyle.Fill;
            l01.Margin = new Padding(10);

            tableLayoutPanel.Controls.Add(l01, 0, 0);
            
            textBox = new TextBox();
            textBox.Anchor = AnchorStyles.Top;
            textBox.Dock = DockStyle.Fill;
            textBox.Margin = new Padding(10);
            textBox.TextChanged += textBox_Changed;

            tableLayoutPanel.Controls.Add(textBox, 0, 1);

            FlowLayoutPanel flowLayoutPanel = new FlowLayoutPanel();

            flowLayoutPanel.Padding = new Padding(0);
            flowLayoutPanel.Margin = new Padding(0);
            flowLayoutPanel.Dock = DockStyle.Fill;
            flowLayoutPanel.FlowDirection = FlowDirection.RightToLeft;

            bOk = new Button();
            bCancel = new Button();

            bOk.Text = "OK";
            bCancel.Text = "Cancel";

            bOk.Click += bOk_Click;
            bCancel.Click += bCancel_Click;

            bOk.Enabled = false;
            bOk.Margin = new Padding(5, 5, 10, 0);
            bCancel.Margin = new Padding(0, 5, 0, 0);

            flowLayoutPanel.Controls.Add(bOk);
            flowLayoutPanel.Controls.Add(bCancel);

            tableLayoutPanel.Controls.Add(flowLayoutPanel, 0, 2);

            textDialog.ShowDialog();
        }

        private void textBox_Changed(object sender, EventArgs e)
        {
            if (textBox.Text.Length == 0)
                bOk.Enabled = false;
            else
                bOk.Enabled = true;
        }

        private void bCancel_Click(object sender, EventArgs e)
        {
            textDialog.Close();
        }

        private void bOk_Click(object sender, EventArgs e)
        {
            word = textBox.Text;
            textDialog.Close();
        }

        private void bClearParam_Click(object sender, EventArgs e)
        {
            lConfParam.Text ="";
            newParam = false;
            bAddEdParam.Enabled = !newParam;
            lVParamLeft.MultiSelect = false;

            fillingLVParamLeft();

            bAddParam.Enabled = false;
            bDelParam.Enabled = false;
            bClearParam.Enabled = false;

            lstParametersForConfig.Clear();

            countPar = 0;
        }

        //Кнопки редактирования
        #region

        private void bDelParam_Click(object sender, EventArgs e)
        {
            try
            {
                lConfParam.Text = lConfParam.Text.Substring(0, lConfParam.Text.LastIndexOf(" + "));
                lstParametersForConfig.RemoveAt(lstParametersForConfig.Count-1);
            }
            catch (Exception) { 
                lConfParam.Text = ""; 
                countPar = 0; 
                lVParamLeft.MultiSelect = false;

                bAddEdParam.Enabled = true;

                bAddParam.Enabled = false;
                bDelParam.Enabled = false;
                bClearParam.Enabled = false;

                lstParametersForConfig.Clear();

                newParam = false;
                fillingLVParamLeft();
            }

            countPar--;
        }

        private void bAddParam_Click(object sender, EventArgs e)
        {
            var items = lVParamLeft.SelectedItems;
            List<Parameter> lst = new List<Parameter>();

            foreach (ListViewItem it in items)
            {
                if (countPar > 0)
                    lConfParam.Text += " + ";

                var para = generalParamsExempl.Where(x => x.Definition.Name == it.Text).FirstOrDefault();
                var para01 = generalParamsType.Where(x => x.Definition.Name == it.Text).FirstOrDefault();

                if (para != null)
                {
                    lConfParam.Text += para.Definition.Name;
                    lstParametersForConfig.Add(para);
                }
                else
                {
                    lstParametersForConfig.Add(para01);
                    lConfParam.Text += para01.Definition.Name;
                }

                countPar++;
            }

            bParamOK.Enabled = true;
            bDelParam.Enabled = true;
        }

        private void bAddEdParam_Click(object sender, EventArgs e)
        {
            var items = lVParamLeft.SelectedItems;
            lVParamLeft.MultiSelect = true;


            try
            {
                lConfParam.Text = items[items.Count - 1].Text + " = ";

                var paramExempl = generalParamsExempl.Where(x => x.Definition.Name == items[items.Count - 1].Text).ToList();
                var paramType = generalParamsType.Where(x => x.Definition.Name == items[items.Count - 1].Text).ToList();
                
                foreach (Parameter p in paramExempl)
                {
                    paramEd = p;
                    lVParamLeft.Items.Remove(new ListViewItem(p.Definition.Name));
                }

                foreach (Parameter p1 in paramType)
                {
                    paramEd = p1;
                    lVParamLeft.Items.Remove(new ListViewItem(p1.Definition.Name));
                }
                if (paramEd != null)
                    newParam = true;
            }
            catch (Exception) { return; }

            fillingLVParamLeft();

            bAddParam.Enabled = true;
            bClearParam.Enabled = true;

            bAddEdParam.Enabled = !newParam;
        }

        private void bParamCancel_Click(object sender, EventArgs e)
        {
            lstParametersForConfig.Clear();
            word = ", ";
            countPar = 0;
            formParam.Close();
        }

        private void bParamBack_Click(object sender, EventArgs e)
        {
            lstParametersForConfig.Clear();
            countPar = 0;
            formParam.Close();
            createFormCategories();
        }

        private void bParamOK_Click(object sender, EventArgs e)
        {
            if (allExempl.First().LookupParameter(paramEd.Definition.Name) != null)
                setParamsForExempl();
            else if (allType.First().LookupParameter(paramEd.Definition.Name) != null)
                setParamsForType();
            else
                MessageBox.Show("Error!");
            formParam.Close();
        }

        #endregion

        List<Element> allExempl = new List<Element>();
        List<Element> allType = new List<Element>();

        private void setParamsForType()
        {
            foreach (Element elem in allType)
            {
                var param = elem.LookupParameter(paramEd.Definition.Name);
                setParamElem(param, elem);
            }
        }

        private void setParamsForExempl()
        {
            int id = -1002050;
            int countT, countE;

            countT = 0;
            countE = 0;

            foreach (Parameter p in lstParametersForConfig)
            {
                if (allType.FirstOrDefault().LookupParameter(p.Definition.Name) != null)
                    countT++;
                if (allExempl.FirstOrDefault().LookupParameter(p.Definition.Name) != null)
                    countE++;
            }

            Trace.WriteLine(countE);
            Trace.WriteLine(countT);


            if (countE > 0 && countT == 0)
            {
                foreach (Element elem in allExempl)
                {
                    var editingParam = elem.get_Parameter((BuiltInParameter)paramEd.Id.IntegerValue);
                    Trace.WriteLine(elem.Name);
                    setParamElem(editingParam, elem);
                }

            }
            else if (countT > 0 && countE == 0)
            {
                foreach (Element elem in allExempl)
                {
                    var typeElem = elem.get_Parameter((BuiltInParameter)id);
                    var editingParam = elem.get_Parameter((BuiltInParameter)paramEd.Id.IntegerValue);

                    Trace.WriteLine(typeElem.Definition.Name);

                    foreach (Element elem02 in allType)
                    {
                        if (typeElem.AsValueString() == elem02.Name)
                            setParamElem(editingParam, elem02);
                    }
                }
            }

            else if (countT > 0 && countE > 0)
            {
                foreach (Element elem in allExempl)
                {
                    int i = 0;
                    string str = "";

                    var editingParam = elem.get_Parameter((BuiltInParameter)paramEd.Id.IntegerValue);
                    var type = elem.get_Parameter((BuiltInParameter)id).AsValueString();

                    foreach (Parameter p in lstParametersForConfig)
                    {
                        var paramElem = elem.LookupParameter(p.Definition.Name);
                        if (paramElem != null)
                        {
                            if (i != 0)
                                str += word + ParameterValueString(paramElem, elem);
                            else
                            {
                                str += ParameterValueString(paramElem, elem) + "";
                                i++;
                            }
                        }
                        else {
                            foreach (Element elem01 in allType)
                            {
                                if (elem01.Name == type)
                                {
                                    var paramType = elem01.LookupParameter(p.Definition.Name);
                                    if (paramType != null)
                                    {
                                        if (i != 0)
                                        {
                                            str += word + ParameterValueString(paramType, elem01);
                                            break;
                                        }
                                        else
                                        {
                                            str += ParameterValueString(paramType, elem01) + "";
                                            i++;
                                            break;
                                        }
                                    }
                                    else break;
                                }
                            }

                        }
                    }
                    setParameter(editingParam, str);
                }
            }
        }

        private void setParamElem(Parameter param, Element elem)
        {
            string str = "";
            int i = 0;

            foreach (Parameter p in lstParametersForConfig)
            {
                var param01 = elem.get_Parameter((BuiltInParameter)p.Id.IntegerValue);

                Trace.WriteLine(param01.AsValueString());

                if (param01 == null)
                    break;

                if (i != 0)
                    str += word + ParameterValueString(param01, elem);
                else
                {
                    str += ParameterValueString(param01, elem) + "";
                    i++;
                }
            }
            setParameter(param, str);
        }

        private string ParameterValueString(Parameter param, Element elem)
        {
                switch (param.StorageType)
                {
                    case StorageType.Double:
                        return param.AsValueString();
                        break;
                    case StorageType.Integer:
                        if (param.Definition.ParameterType == ParameterType.YesNo)
                        {
                            if (param.AsInteger() == 0)
                                return "Нет";
                            else
                                return "Да";
                        }
                        return param.AsValueString();
                        break;
                    case StorageType.String:
                        return param.AsString();
                        break;
                    case StorageType.ElementId:
                        Document doc01 = param.Element.Document;
                        string znachenie = "<null>";
                        Element el = doc01.GetElement(param.AsElementId());
                        if (el != null) 
                            znachenie = el.Name;
                        else
                        {
                            znachenie = param.AsValueString();
                        }
                    if (znachenie.Length == 0)
                        znachenie = "<null>";

                    return znachenie;
                        break;
                    default:
                        return "<null>";
                        break;
                }
            
        }

        private void setParameter(Parameter par, string str)
        {
            if (par != null)
            {
                switch (par.StorageType)
                {
                    case StorageType.Double:
                        try
                        {
                            par.Set(double.Parse(str));
                        }
                        catch
                        { break; }
                        break;
                    case StorageType.String:
                        try
                        {
                            par.Set(str);
                        }
                        catch
                        { break; }
                        break;
                    case StorageType.Integer:
                        try
                        {
                            par.Set(int.Parse(str));
                        }
                        catch
                        { break; }
                        break;
                    case StorageType.ElementId:
                        try
                        {
                            par.SetValueString(str);
                        }
                        catch
                        { break; }
                        break;
                    default:
                        break;
                }
            }
        }


        private void createFLParamRightPanelButtons()
        {
            fLParamRightPanelButtons = new FlowLayoutPanel();
            fLParamRightPanelButtons.Dock = DockStyle.Fill;
            fLParamRightPanelButtons.FlowDirection = FlowDirection.TopDown;
            fLParamRightPanelButtons.Margin = new Padding(0);

            setLayoutFLParamRightPanelButtons();
        }

        private void setLayoutFLParamRightPanelButtons()
        {
            fLParamRightPanelButtons.SuspendLayout();

            fLParamRightPanelButtons.Controls.Add(bAddEdParam);
            fLParamRightPanelButtons.Controls.Add(bAddParam);
            fLParamRightPanelButtons.Controls.Add(bDelParam);
            fLParamRightPanelButtons.Controls.Add(bClearParam);
            fLParamRightPanelButtons.Controls.Add(bSetWordParam);

            fLParamRightPanelButtons.ResumeLayout();
            fLParamRightPanelButtons.Refresh();
        }
        #endregion

        private Categories getAllCategories() { return document.Settings.Categories; }

        private IList<Element> getFamilies(Category category) {
            int id = Int32.Parse(category.Id.ToString());
            return new FilteredElementCollector(document).OfCategory((BuiltInCategory)id).WhereElementIsNotElementType().ToElements();
        }

        private IList<Element> getTypeFamilies(Category category)
        {
            int id = Int32.Parse(category.Id.ToString());
            return new FilteredElementCollector(document).OfCategory((BuiltInCategory)id).WhereElementIsElementType().ToElements();
        }

        private void collectGenParametersExempl()
        {
            int i = 0;

            List<int> fpId = new List<int>();
            List<int> generId = new List<int>();

            foreach (List<ParameterSet> lst in allSelectedParamsExempl)
            {
                foreach (ParameterSet par in lst)
                {
                    if (i == 0)
                    {
                        fpId = getParamId(par);
                        generId = fpId;
                    }
                    else
                        generId = (List<int>)generId.Intersect(getParamId(par));
                }
            }

            if (gEx == 0)
            {
                fParamIdExempl = generId;
                generParamIdExempl = fParamIdExempl;
                gEx++;
            }
            else
                generParamIdExempl = generParamIdExempl.Intersect(generId).ToList();
        }

        private void collectGenParametersType() {
            int i = 0;

            List<int> fpId = new List<int>();
            List<int> generId = new List<int>();

            foreach (List<ParameterSet> lst in allSelectedParamsType)
            {
                foreach (ParameterSet par in lst)
                {
                    if (i == 0)
                    {
                        fpId = getParamId(par);
                        generId = fpId;
                    }
                    else
                        generId = (List<int>)generId.Intersect(getParamId(par));
                }
            }

            if (gType == 0)
            {
                fParamIdType = generId;
                generParamIdType = fParamIdType;
                gType++;
            }
            else
                generParamIdType = generParamIdType.Intersect(generId).ToList();
        }

        private void getGeneralParamExempl()
        {
            foreach (Category elem in checkedCategory)
            {
                try
                {
                    Element fam = getFamilies(elem).First();
                    foreach (Parameter p in fam.Parameters)
                        foreach (int pId in generParamIdExempl)
                        {
                            if (pId == p.Id.IntegerValue)
                                generalParamsExempl.Add(p);
                        }
                }
                catch(Exception) { }
                
            }
        }

        private void getGeneralParamType()
        {
            foreach (Category elem in checkedCategory)
            {
                try
                {
                    Element fam = getTypeFamilies(elem).First();
                    foreach (Parameter p in fam.Parameters)
                        foreach (int pId in generParamIdType)
                            if (pId == p.Id.IntegerValue)
                                generalParamsType.Add(p);
                }
                catch (Exception) { }
            }
        }

        private List<int> getParamId(ParameterSet param) {

            List<int> pId = new List<int>();
            foreach (Parameter p in param)
                pId.Add(p.Id.IntegerValue);

            return pId;
        }


        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDoc = commandData.Application.ActiveUIDocument;
            document = UIDoc.Document;

            Transaction transaction = new Transaction(document, "Configure Param");
            transaction.Start();

            createFormCategories();

            transaction.Commit();

            return Result.Succeeded;
        }

    }
}
