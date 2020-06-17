using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;

namespace TitleListGenerator
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }



        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (checkValidateData())
            {
                switch (comboType.SelectedIndex)
                {
                    case 0:
                        generateSemesterTitle();
                        break;
                    default:
                        break;
                }
            }
        }

        private Boolean checkValidateData()
        {
            if (comboType.SelectedIndex == -1)
            {
                MessageBox.Show("Не выбран тип работы",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                return false;
            }
            if (textCathedra.Text.Trim().Length == 0)
            {
                MessageBox.Show("Не заполнено поле Направление подготовки",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                return false;
            }
            if (textDiscipline.Text.Trim().Length == 0)
            {
                MessageBox.Show("Не заполнено поле Дисциплина",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                return false;
            }
            if (textFaculty.Text.Trim().Length == 0)
            {
                MessageBox.Show("Не заполнено поле Факультет",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                return false;
            }
            if (textGroup.Text.Trim().Length == 0)
            {
                MessageBox.Show("Не заполнено поле Группа",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                return false;
            }
            if (textName.Text.Trim().Length == 0)
            {
                MessageBox.Show("Не заполнено поле ФИО",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                return false;
            }
            if (textProfessor.Text.Trim().Length == 0)
            {
                MessageBox.Show("Не заполнено поле Преподаватель",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                return false;
            }
            if (textSubject.Text.Trim().Length == 0)
            {
                MessageBox.Show("Не заполнено поле Предмет",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                return false;
            }
            return true;
        }

        private void generateSemesterTitle()
        {
            String path = Environment.CurrentDirectory + "\\SemesterTitle.dotx";
            Word.Application app = new Word.Application();
            // Создаём объект документа
            Word.Document doc = app.Documents.Add(path);
            doc.Activate();

            doc.Bookmarks["NAME"].Range.Text = textName.Text;
            doc.Bookmarks["DIRECTION"].Range.Text = textCathedra.Text;
            doc.Bookmarks["FACULTY"].Range.Text = textFaculty.Text;
            doc.Bookmarks["GROUP"].Range.Text = textGroup.Text;
            doc.Bookmarks["PROFESSOR"].Range.Text = textProfessor.Text;
            doc.Bookmarks["SUBJECT"].Range.Text = textSubject.Text;
            doc.Bookmarks["SUBJECT_LESSONS"].Range.Text = textDiscipline.Text;
            doc.Bookmarks["YEAR"].Range.Text = numYear.Value.ToString();


            doc.SaveAs(FileName: Environment.CurrentDirectory + "\\" + textName.Text + "_" + textDiscipline.Text + ".docx");
            doc.Close();

            MessageBox.Show("Файл " + textName.Text + "_" + textDiscipline.Text + ".docx успешно сгенерирован!",
                                "Успех",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
        }
    }
}
