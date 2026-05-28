using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Infinium
{
    public partial class AddProjectForm : Form
    {
        private const int eHide = 2;
        private const int eShow = 1;
        private const int eClose = 3;

        private int FormEvent;

        public bool Canceled;

        private bool bEdit;

        private Form TopForm;

        public bool bProposition;

        private TreeView TempMembersList;

        private InfiniumProjects InfiniumProjects;

        public int ProjectID = -1;

        public AddProjectForm(ref InfiniumProjects tInfiniumProjects, ref Form tTopForm)
        {
            InitializeComponent();

            TopForm = tTopForm;
            InfiniumProjects = tInfiniumProjects;

            SetMembersTree();
        }

        public AddProjectForm(ref InfiniumProjects tInfiniumProjects, ref Form tTopForm, int ProjectID)//edit
        {
            InitializeComponent();

            TopForm = tTopForm;
            InfiniumProjects = tInfiniumProjects;

            bEdit = true;

            ProjectNameTextEdit.Text = InfiniumProjects.ProjectsDataTable.Select("ProjectID = " + ProjectID)[0]["ProjectName"].ToString();
            ProjectDescriptionRichTextBox.Text = InfiniumProjects.ProjectsDataTable.Select("ProjectID = " + ProjectID)[0]["ProjectDescription"].ToString();
            ProposCheckBox.Checked = Convert.ToBoolean(InfiniumProjects.ProjectsDataTable.Select("ProjectID = " + ProjectID)[0]["IsProposition"]);

            AllUsersNotifyCheckBox.Enabled = false;
            AllUsersNotifyCheckBox.StateCommon.ShortText.Color1 = Color.LightGray;

            SetMembersTree();
            CheckMembersTree();
        }

        public void CheckMembersTree()
        {
            if (InfiniumProjects.CurrentProjectDepartmentsDataTable.Rows.Count > 0)
            {
                foreach (DataRow Row in InfiniumProjects.CurrentProjectDepartmentsDataTable.Rows)
                {
                    foreach (TreeNode Node in MembersTree.Nodes)
                    {
                        if (Node.Text == Row["DepartmentName"].ToString())
                            Node.Checked = true;
                    }

                    foreach (TreeNode Node in TempMembersList.Nodes)
                    {
                        if (Node.Text == Row["DepartmentName"].ToString())
                        {
                            Node.Checked = true;

                            foreach (TreeNode SubNode in Node.Nodes)
                            {
                                SubNode.Checked = true;
                            }
                        }
                    }
                }
            }

            if (InfiniumProjects.CurrentProjectUsersDataTable.Rows.Count > 0)
            {
                foreach (DataRow Row in InfiniumProjects.CurrentProjectUsersDataTable.Rows)
                {
                    foreach (TreeNode Node in MembersTree.Nodes)
                    {
                        bool Expand = false;

                        foreach (TreeNode SubNode in Node.Nodes)
                            if (SubNode.Text == Row["Name"].ToString())
                            {
                                SubNode.Checked = true;
                                Expand = true;
                            }

                        if (Expand)
                            Node.Expand();
                    }


                    foreach (TreeNode Node in TempMembersList.Nodes)
                    {
                        foreach (TreeNode SubNode in Node.Nodes)
                            if (SubNode.Text == Row["Name"].ToString())
                                SubNode.Checked = true;
                    }
                }
            }
        }

        public void SetMembersTree()
        {
            TempMembersList = new TreeView();

            for (int i = 0; i < InfiniumProjects.DepartmentsDataTable.Rows.Count; i++)
            {
                DataRow[] URows = InfiniumProjects.UsersDataTable.Select("DepartmentID = " +
                                            InfiniumProjects.DepartmentsDataTable.Rows[i]["DepartmentID"]);

                MembersTree.Nodes.Add(InfiniumProjects.DepartmentsDataTable.Rows[i]["DepartmentName"].ToString());
                TempMembersList.Nodes.Add(InfiniumProjects.DepartmentsDataTable.Rows[i]["DepartmentName"].ToString());

                for (int j = 0; j < URows.Count(); j++)
                {
                    MembersTree.Nodes[i].Nodes.Add(URows[j]["Name"].ToString());
                    TempMembersList.Nodes[i].Nodes.Add(URows[j]["Name"].ToString());
                }
            }
        }

        public bool CompareMembersTrees()
        {
            for (int i = 0; i < MembersTree.Nodes.Count; i++)
            {
                if (TempMembersList.Nodes[i].Checked != MembersTree.Nodes[i].Checked)
                    return true;

                for (int j = 0; j < MembersTree.Nodes[i].GetNodeCount(true); j++)
                {
                    if (TempMembersList.Nodes[i].Nodes[j].Checked != MembersTree.Nodes[i].Nodes[j].Checked)
                        return true;
                }
            }

            return false;
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == 6 && m.WParam.ToInt32() == 1)
            {
                if (TopForm != null)
                    TopForm.Activate();
            }
        }

        private void AnimateTimer_Tick(object sender, EventArgs e)
        {
            if (!DatabaseConfigsManager.Animation)
            {
                Opacity = 1;

                if (FormEvent == eClose || FormEvent == eHide)
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {
                        Close();
                    }

                    if (FormEvent == eHide)
                    {
                        Hide();
                    }

                    return;
                }

                if (FormEvent == eShow)
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                    return;
                }

            }

            if (FormEvent == eClose || FormEvent == eHide)
            {
                if (Convert.ToDecimal(Opacity) != Convert.ToDecimal(0.00))
                    Opacity = Convert.ToDouble(Convert.ToDecimal(Opacity) - Convert.ToDecimal(0.05));
                else
                {
                    AnimateTimer.Enabled = false;

                    if (FormEvent == eClose)
                    {

                        Close();
                    }

                    if (FormEvent == eHide)
                    {
                        Hide();
                    }
                }

                return;
            }


            if (FormEvent == eShow || FormEvent == eShow)
            {
                if (Opacity != 1)
                    Opacity += 0.05;
                else
                {
                    AnimateTimer.Enabled = false;
                    SplashForm.CloseS = true;
                }
            }
        }

        private void AddNewsForm_Shown(object sender, EventArgs e)
        {
            FormEvent = eShow;
            AnimateTimer.Enabled = true;
        }

        private void CancelNewsButton_Click(object sender, EventArgs e)
        {
            Canceled = true;

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void MembersTree_AfterCheck(object sender, TreeViewEventArgs e)
        {
            for (int i = 0; i < e.Node.Nodes.Count; i++)
            {
                e.Node.Nodes[i].Checked = e.Node.Checked;
            }
        }

        private void OKNewsButton_Click(object sender, EventArgs e)
        {
            if (ProjectNameTextEdit.Text.Length < 1)
            {
                InfiniumTips.ShowTip(this, 50, 85, "Введите название проекта", 2300);
                return;
            }

            if (ProjectDescriptionRichTextBox.Text.Length < 1)
            {
                InfiniumTips.ShowTip(this, 50, 85, "Введите описание проекта", 2300);
                return;
            }

            bool bChecked = false;

            bProposition = ProposCheckBox.Checked;

            foreach (TreeNode Node in MembersTree.Nodes)
            {
                if (Node.Checked)
                {
                    bChecked = true;
                    break;
                }

                foreach (TreeNode SubNode in Node.Nodes)
                {
                    if (SubNode.Checked)
                    {
                        bChecked = true;
                        break;
                    }
                }
            }

            if (!bChecked)
            {
                InfiniumTips.ShowTip(this, 50, 85, "В проекте должен быть хотя бы один участник (сотрудник или целый отдел)", 3300);
                return;
            }

            CreateLabel.Visible = true;

            Application.DoEvents();

            bool bNeedMembersUpdate = CompareMembersTrees();

            if (!bEdit)
            {
                ProjectID = InfiniumProjects.AddProject(MembersTree, ProjectNameTextEdit.Text, ProjectDescriptionRichTextBox.Text,
                                        Security.CurrentUserID, ProposCheckBox.Checked);

                InfiniumProjects.AddProjectSubscribe(ProjectID, MembersTree, AllUsersNotifyCheckBox.Checked, ProposCheckBox.Checked);
            }
            else
                InfiniumProjects.EditProject(MembersTree, bNeedMembersUpdate, ProjectNameTextEdit.Text, ProjectDescriptionRichTextBox.Text,
                                             ProjectID, ProposCheckBox.Checked);

            FormEvent = eClose;
            AnimateTimer.Enabled = true;
        }

        private void ProjectNameTextEdit_TextChanged(object sender, EventArgs e)
        {
            SymbolCountText.Text = (ProjectNameTextEdit.MaxLength - ProjectNameTextEdit.Text.Length) + " символов осталось";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(MembersTree.Nodes[0].Text);
        }

        private void ProjectDescriptionRichTextBox_TextChanged(object sender, EventArgs e)
        {
            DescCountLabel.Text = (ProjectDescriptionRichTextBox.MaxLength - ProjectDescriptionRichTextBox.Text.Length) + " символов осталось";
        }
    }
}
