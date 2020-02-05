
using MaterialSkin.Controls;
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

namespace Shablonizator_2._0
{
    public partial class Form1 : MaterialForm
    {
        public Form1()
        {
            InitializeComponent();
            MaterialSkin.MaterialSkinManager SkinManager = MaterialSkin.MaterialSkinManager.Instance;
            SkinManager.AddFormToManage(this);
            SkinManager.Theme = MaterialSkin.MaterialSkinManager.Themes.LIGHT;//цвет окна 
            var Panel1 = MaterialSkin.Primary.Blue900;
            var Panel2 = MaterialSkin.Primary.Blue800;
            var AnyPareamet = MaterialSkin.Accent.Blue700;
            var Text = MaterialSkin.TextShade.WHITE;
            SkinManager.ColorScheme = new MaterialSkin.ColorScheme(Panel2, Panel1, MaterialSkin.Primary.Purple700, AnyPareamet, Text);
            //вторая панель и цвет кнопки,верхняя панель, хз,цвет чекбоксов и тд., шрифт 
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var Object = tabPage1;
            for (int i = 1; i <= materialTabControl1.TabCount-1;i++ )
            {
                switch (i)
                {
                    case 1: Object = tabPage1; break;
                    case 2: Object = tabPage2; break;
                    case 3: Object = tabPage3; break;
                    case 4: Object = tabPage4; break;
                    case 5: Object = tabPage5; break;
                }
                foreach (MaterialSingleLineTextField TextBox in Object.Controls.OfType<MaterialSingleLineTextField>())
                {
                        TextBox.Text = "";
                }
            }
            int Count = 4;
            for (int j = 1; j <= Count; j++)
            {
                switch (j)
                {
                    case 1: materialSingleLineTextField49.Text = Properties.Settings.Default.Pattern1; break;
                    case 2: materialSingleLineTextField50.Text = Properties.Settings.Default.Pattern2; break;
                    case 3: materialSingleLineTextField51.Text = Properties.Settings.Default.Pattern3; break;
                    case 4: materialSingleLineTextField52.Text = Properties.Settings.Default.Pattern4; break;
                }
            }

        }

        Point point = new Point();
        public int Index;
        public void SelectGrid(int GridObject, int GeridSelectedRow, int Action)
        {
            try
            {
                var Grid = bunifuCustomDataGrid1;
                var TextBox = materialSingleLineTextField1;
                var TextBox2 = materialSingleLineTextField1;
                string text;

                switch (GridObject)
                {
                    case 1: Grid = bunifuCustomDataGrid1; TextBox = materialSingleLineTextField6; TextBox2 = null; break;
                    case 2: Grid = bunifuCustomDataGrid2; TextBox = materialSingleLineTextField7; TextBox2 = null; break;
                    case 3: Grid = bunifuCustomDataGrid3; TextBox = materialSingleLineTextField8; TextBox2 = null; break;
                    case 4: Grid = bunifuCustomDataGrid4; TextBox = materialSingleLineTextField9; TextBox2 = null; break;
                    case 5: Grid = bunifuCustomDataGrid5; TextBox = materialSingleLineTextField19; TextBox2 = null; break;
                    case 6: Grid = bunifuCustomDataGrid6; TextBox = materialSingleLineTextField20; TextBox2 = null; break;
                    case 7: Grid = bunifuCustomDataGrid7; TextBox = materialSingleLineTextField30; TextBox2 = null; break;
                    case 8: Grid = bunifuCustomDataGrid8; TextBox = materialSingleLineTextField29; TextBox2 = null; break;
                    case 9: Grid = bunifuCustomDataGrid9; TextBox = materialSingleLineTextField28; TextBox2 = null; break;
                    case 10: Grid = bunifuCustomDataGrid10; TextBox = materialSingleLineTextField45; TextBox2 = materialSingleLineTextField46; break; 
                    case 11: Grid = bunifuCustomDataGrid11; TextBox = materialSingleLineTextField44; TextBox2 = null; break;
                }

                if (TextBox2 != null)
                {
                    text = TextBox.Text + " " + TextBox2.Text;
                }
                else
                {
                    text = TextBox.Text;
                }


                switch (Action)
                {
                    case 1: Grid.Rows.Add(text); break;
                    case 2: Grid.Rows.RemoveAt(GeridSelectedRow); break;
                    case 3: break;
                    case 4: break;
                }
            }
            catch
            {

            }
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            materialSingleLineTextField1.Text = monthCalendar1.SelectionStart.Date.ToShortDateString();
        }

        private void materialFlatButton1_Click(object sender, EventArgs e)
        {
            if (materialSingleLineTextField6.Text.Length != 0)
            {
                SelectGrid(1, 0, 1);
            }
            else
            {
                MessageBox.Show("Поле ввода не содержит записей!");
            }
        }

        private void materialFlatButton4_Click(object sender, EventArgs e)
        {
            if (materialSingleLineTextField7.Text.Length != 0)
            {
                SelectGrid(2, 0, 1);
            }
            else
            {
                MessageBox.Show("Поле ввода не содержит записей!");
            }
        }

        private void materialFlatButton6_Click(object sender, EventArgs e)
        {
            if (materialSingleLineTextField8.Text.Length != 0)
            {
                SelectGrid(3, 0, 1);
            }
            else
            {
                MessageBox.Show("Поле ввода не содержит записей!");
            }
        }

        private void materialFlatButton8_Click(object sender, EventArgs e)
        {
            if (materialSingleLineTextField9.Text.Length != 0)
            {
                SelectGrid(4, 0, 1);
            }
            else
            {
                MessageBox.Show("Поле ввода не содержит записей!");
            }
        }

        public int GridObjectPublic;
        public void LeftClick(int GridObject)
        {
            if (Convert.ToBoolean(MouseButtons.Left))
            {
                var Grid = bunifuCustomDataGrid1;
                switch (GridObject)
                {
                    case 1: Grid = bunifuCustomDataGrid1; break;
                    case 2: Grid = bunifuCustomDataGrid2; break;
                    case 3: Grid = bunifuCustomDataGrid3; break;
                    case 4: Grid = bunifuCustomDataGrid4; break;
                    case 5: Grid = bunifuCustomDataGrid5; break;
                    case 6: Grid = bunifuCustomDataGrid6; break;
                    case 7: Grid = bunifuCustomDataGrid7; break;
                    case 8: Grid = bunifuCustomDataGrid8; break;
                    case 9: Grid = bunifuCustomDataGrid9; break;
                    case 10: Grid = bunifuCustomDataGrid10; break;
                    case 11: Grid = bunifuCustomDataGrid11; break;
                }
                GridObjectPublic = GridObject;
                point = new Point(50, 50);
                Grid.ContextMenuStrip = materialContextMenuStrip1;
                materialContextMenuStrip1.Show(Grid, new Point(point.X, point.Y));
                Grid.ContextMenuStrip = null;
            }
        }
        private void bunifuCustomDataGrid1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LeftClick(1);
            Index = e.RowIndex;
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectGrid(GridObjectPublic, 0, 2);
        }

        private void bunifuCustomDataGrid2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LeftClick(2);
            Index = e.RowIndex;
        }

        private void bunifuCustomDataGrid3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LeftClick(3);
            Index = e.RowIndex;
        }

        private void bunifuCustomDataGrid4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LeftClick(4);
            Index = e.RowIndex;
        }

        public void CountTextBoxforForm(int Action)
        {
            var Page = tabPage1;
            var Grid = bunifuCustomDataGrid1;
            int Action1 = 0;
            int Action2 = 0;
            int startGrid = 0;
            int endGrid = 0;
            switch (Action)
            {
                case 1: Page = tabPage1; break;
                case 2: Page = tabPage2; break;
                case 3: Page = tabPage3; break;
                case 4: Page = tabPage4; break;
            }
            foreach (MaterialSingleLineTextField TextBox in Page.Controls.OfType<MaterialSingleLineTextField>())
            {
                if (TextBox.Text.Length == 0)
                {
                    Action1++;
                }
            }
            if (Action1 == 0)
            {
                switch (Action)
                {
                    case 1:
                        {
                            startGrid = 1; endGrid = 4;
                        } break;
                    case 2:
                        {
                            startGrid = 5; endGrid = 6;
                        } break;
                    case 3:
                        {
                            startGrid = 7; endGrid = 9;
                        } break;
                    case 4:
                        {
                            startGrid = 10; endGrid = 11;
                        } break;
                }

                for (int i = startGrid; i <= endGrid; i++)
                {
                    switch (i)
                    {
                        case 1: Grid = bunifuCustomDataGrid1; break;
                        case 2: Grid = bunifuCustomDataGrid2; break;
                        case 3: Grid = bunifuCustomDataGrid3; break;
                        case 4: Grid = bunifuCustomDataGrid4; break;
                        case 5: Grid = bunifuCustomDataGrid5; break;
                        case 6: Grid = bunifuCustomDataGrid6; break;
                        case 7: Grid = bunifuCustomDataGrid7; break;
                        case 8: Grid = bunifuCustomDataGrid8; break;
                        case 9: Grid = bunifuCustomDataGrid9; break;
                        case 10: Grid = bunifuCustomDataGrid10; break;
                        case 11: Grid = bunifuCustomDataGrid11; break;
                    }
                    if (Grid.Rows.Count == 0)
                    {
                        Action2++;
                    }
                }

                if (Action2 != 0)
                {
                    MessageBox.Show("Некоторые списки не содержат записей!");
                }
                else
                {
                    Export(Action);
                }
            }
            else
            {
                MessageBox.Show("Некоторые поял не заполены!");
            }

        }
        public string Path = "";
        public void Export(int Action)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            var wordApp = new Word.Application();//переменная для word
            openFileDialog1.Filter = "Документы (*.docx)|*.docx|Документы (*.doc)|*.doc";
            openFileDialog1.FileName = "";
            string str = "";
            string strReplace = "";
            var MatText = materialSingleLineTextField1;
            var Grid = bunifuCustomDataGrid1;
            int TableEndCount = 0;
            int TableStartCount = 0;
            int TextBoxStartCount = 0;
            int TextBoxEndCount = 0;
            int ActiveTable = 0;

            switch (Action)
            {
                case 1: Path = materialSingleLineTextField49.Text; break;
                case 2: Path = materialSingleLineTextField50.Text; break;
                case 3: Path = materialSingleLineTextField51.Text; break;
                case 4: Path = materialSingleLineTextField52.Text; break;
            }
            if (Path.Length == 0)
            {
                MessageBox.Show("Файл не найден, пожалуйста укажите его вручную.");
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Path = openFileDialog1.FileName;
                }
            }
            else
            {

                var wordDocument = wordApp.Documents.Open(Path);
                wordApp.Visible = false;//word скрыт

                switch (Action)
                {
                    case 1:
                        {
                            TextBoxStartCount = 1;
                            TextBoxEndCount = 10;
                           
                            TableStartCount = 1;
                            TableEndCount = 4;   
                        } break;
                    case 2:
                        {
                            TextBoxStartCount = 11;
                            TextBoxEndCount = 20;

                            TableStartCount = 5;
                            TableEndCount = 6;   
                        } break;
                    case 3:
                        {
                            TextBoxStartCount = 21;
                            TextBoxEndCount = 30;

                            TableStartCount = 7;
                            TableEndCount = 9; 
                        } break;
                    case 4:
                        {
                            TextBoxStartCount = 31;
                            TextBoxEndCount = 37;

                            TableStartCount = 10;
                            TableEndCount = 11; 
                        } break;
                }

                for (int i = TextBoxStartCount; i <= TextBoxEndCount; i++)
                {
                    switch (i)
                    {
                        case 1: MatText = materialSingleLineTextField1; str = MatText.Text; strReplace = "<Дата>"; break;
                        case 2: MatText = materialSingleLineTextField2; str = MatText.Text; strReplace = "<Номер>"; break;
                        case 3: MatText = materialSingleLineTextField3; str = MatText.Text; strReplace = "<Председатель0>"; break;
                        case 4: MatText = materialSingleLineTextField4; str = MatText.Text; strReplace = "<Заместитель>"; break;
                        case 5: MatText = materialSingleLineTextField5; str = MatText.Text; strReplace = "<Секретарь0>"; break;
                        case 6: MatText = materialSingleLineTextField10; str = MatText.Text; strReplace = "<Председатель1>"; break;
                        case 7: MatText = materialSingleLineTextField11; str = MatText.Text; strReplace = "<Председатель2>"; break;
                        case 8: MatText = materialSingleLineTextField12; str = MatText.Text; strReplace = "<ПредседательЦМК>"; break;
                        case 9: MatText = materialSingleLineTextField13; str = MatText.Text; strReplace = "<Преподаватель>"; break;
                        case 10: MatText = materialSingleLineTextField14; str = MatText.Text; strReplace = "<Секретарь1>"; break;

                        case 11: MatText = materialSingleLineTextField16; str = MatText.Text; strReplace = "<Дисциплина0>"; break;
                        case 12: MatText = materialSingleLineTextField17; str = MatText.Text; strReplace = "<Председатель0>"; break;
                        case 13: MatText = materialSingleLineTextField15; str = MatText.Text; strReplace = "<Дата>"; break;
                        case 14: MatText = materialSingleLineTextField18; str = MatText.Text; strReplace = "<Группы>"; break;
                        case 15: MatText = materialSingleLineTextField16; str = MatText.Text; strReplace = "<Дисциплина1>"; break;
                        case 16: MatText = materialSingleLineTextField25; str = MatText.Text; strReplace = "<Председатель1>"; break;
                        case 17: MatText = materialSingleLineTextField23; str = MatText.Text; strReplace = "<ПредседательЦМК>"; break;
                        case 18: MatText = materialSingleLineTextField22; str = MatText.Text; strReplace = "<Преподаватель1>"; break;
                        case 19: MatText = materialSingleLineTextField24; str = MatText.Text; strReplace = "<Заведущая0>"; break;
                        case 20: MatText = materialSingleLineTextField21; str = MatText.Text; strReplace = "<Заведущая1>"; break;

                        case 21: MatText = materialSingleLineTextField26; str = MatText.Text; strReplace = "<Заведущая1>"; break;
                        case 22: MatText = materialSingleLineTextField27; str = MatText.Text; strReplace = "<Председатель0>"; break;
                        case 23: MatText = materialSingleLineTextField31; str = MatText.Text; strReplace = "<Протокол0>"; break;
                        case 24: MatText = materialSingleLineTextField32; str = MatText.Text; strReplace = "<Протокол1>"; break;
                        case 25: MatText = materialSingleLineTextField33; str = MatText.Text; strReplace = "<Группы0>"; break;
                        case 26: MatText = materialSingleLineTextField34; str = MatText.Text; strReplace = "<Модуль0>"; break;
                        case 27: MatText = materialSingleLineTextField35; str = MatText.Text; strReplace = "<Группы1>"; break;
                        case 28: MatText = materialSingleLineTextField36; str = MatText.Text; strReplace = "<Модуль1>"; break;
                        case 29: MatText = materialSingleLineTextField37; str = MatText.Text; strReplace = "<Место0>"; break;
                        case 30: MatText = materialSingleLineTextField37; str = MatText.Text; strReplace = "<Зам>"; break;

                        case 31: MatText = materialSingleLineTextField40; str = MatText.Text; strReplace = "<Дата>"; break;
                        case 32: MatText = materialSingleLineTextField39; str = MatText.Text; strReplace = "<Номер>"; break;
                        case 33: MatText = materialSingleLineTextField41; str = MatText.Text; strReplace = "<Секретарь0>"; break;
                        case 34: MatText = materialSingleLineTextField42; str = MatText.Text; strReplace = "<Председатель0>"; break;
                        case 35: MatText = materialSingleLineTextField43; str = MatText.Text; strReplace = "<Заместитель>"; break;
                        case 36: MatText = materialSingleLineTextField48; str = MatText.Text; strReplace = "<Председатель1>"; break;
                        case 37: MatText = materialSingleLineTextField47; str = MatText.Text; strReplace = "<Секретарь1>"; break;

                    }
                    ReplaceWordsStub(strReplace, str, wordDocument); 
                }
                try
                {
                    for (int Table = TableStartCount; Table <= TableEndCount; Table++)
                    {
                        switch (Table)
                        {
                            case 1: Grid = bunifuCustomDataGrid1; break;
                            case 2: Grid = bunifuCustomDataGrid2; break;
                            case 3: Grid = bunifuCustomDataGrid3; break;
                            case 4: Grid = bunifuCustomDataGrid4; break;
                            case 5: Grid = bunifuCustomDataGrid5; break;
                            case 6: Grid = bunifuCustomDataGrid6; break;
                            case 7: Grid = bunifuCustomDataGrid7; break;
                            case 8: Grid = bunifuCustomDataGrid8; break;
                            case 9: Grid = bunifuCustomDataGrid9; break;
                            case 10: Grid = bunifuCustomDataGrid10; break;
                            case 11: Grid = bunifuCustomDataGrid11; break;
                        }
                        ActiveTable++;
                        Word.Table _table = wordDocument.Tables[ActiveTable];
                        for (int iTable = 0; iTable <= Grid.Rows.Count - 1; iTable++)
                        {
                            _table.Cell(iTable + 1, 1).Range.Text = Grid.Rows[iTable].Cells[0].Value.ToString();
                            _table.Rows.Add();
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("ПРОВЕРЬТЕ ТАБЛИЦЫ ДОКУМЕНТА, ОШИБКА!");
                }
                

                try
                {
                        wordDocument.SaveAs("");//сохроняем наш документ
                }
                catch
                {

                }
                finally
                {
                    if (materialCheckBox1.Checked)
                    {
                        wordDocument.PrintOut();
                    }
                    wordDocument.Close();//закрываем документ
                    wordApp.Quit();

                    MessageBox.Show("Документ сохранен.");
                }
            }
        }
        private void ReplaceWordsStub(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;//перменная для хранения данных документа
            range.Find.ClearFormatting();//метод сброса всех натсроек текста
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);//находим ключевые слова и заменяем их
        }
        private void materialFlatButton10_Click(object sender, EventArgs e)
        {
            CountTextBoxforForm(1);
        }

        private void materialTabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void materialLabel6_Click(object sender, EventArgs e)
        {

        }

        private void materialSingleLineTextField6_Click(object sender, EventArgs e)
        {

        }

        private void materialSingleLineTextField7_Click(object sender, EventArgs e)
        {

        }

        private void materialLabel7_Click(object sender, EventArgs e)
        {

        }

        private void materialFlatButton3_Click(object sender, EventArgs e)
        {
            if (materialSingleLineTextField20.Text.Length != 0)
            {
                SelectGrid(6, 0, 1);
            }
            else
            {
                MessageBox.Show("Поле ввода не содержит записей!");
            }
        }

        private void materialFlatButton2_Click(object sender, EventArgs e)
        {
            if (materialSingleLineTextField19.Text.Length != 0)
            {
                SelectGrid(5, 0, 1);
            }
            else
            {
                MessageBox.Show("Поле ввода не содержит записей!");
            }
        }

        private void materialFlatButton11_Click(object sender, EventArgs e)
        {
            if (materialSingleLineTextField28.Text.Length != 0)
            {
                SelectGrid(9, 0, 1);
            }
            else
            {
                MessageBox.Show("Поле ввода не содержит записей!");
            }
        }

        private void materialFlatButton12_Click(object sender, EventArgs e)
        {
            if (materialSingleLineTextField29.Text.Length != 0)
            {
                SelectGrid(8, 0, 1);
            }
            else
            {
                MessageBox.Show("Поле ввода не содержит записей!");
            }
        }

        private void materialFlatButton13_Click(object sender, EventArgs e)
        {
            if (materialSingleLineTextField30.Text.Length != 0)
            {
                SelectGrid(7, 0, 1);
            }
            else
            {
                MessageBox.Show("Поле ввода не содержит записей!");
            }
        }

        private void materialFlatButton17_Click(object sender, EventArgs e)
        {
            if (materialSingleLineTextField45.Text.Length != 0 && materialSingleLineTextField46.Text.Length != 0)
            {
                SelectGrid(10, 0, 1);
            }
            else
            {
                MessageBox.Show("Поле ввода не содержит записей!");
            }
        }

        private void materialFlatButton16_Click(object sender, EventArgs e)
        {
            if (materialSingleLineTextField44.Text.Length != 0)
            {
                SelectGrid(11, 0, 1);
            }
            else
            {
                MessageBox.Show("Поле ввода не содержит записей!");
            }
        }

        private void bunifuCustomDataGrid5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LeftClick(5);
            Index = e.RowIndex;
        }

        private void bunifuCustomDataGrid6_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LeftClick(6);
            Index = e.RowIndex;
        }

        private void bunifuCustomDataGrid9_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LeftClick(9);
            Index = e.RowIndex;
        }

        private void bunifuCustomDataGrid8_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LeftClick(8);
            Index = e.RowIndex;
        }

        private void bunifuCustomDataGrid7_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LeftClick(7);
            Index = e.RowIndex;
        }

        private void bunifuCustomDataGrid11_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LeftClick(11);
            Index = e.RowIndex;
        }

        private void bunifuCustomDataGrid10_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LeftClick(10);
            Index = e.RowIndex;
        }

        private void materialFlatButton5_Click(object sender, EventArgs e)
        {
            CountTextBoxforForm(2);
        }

        private void materialFlatButton14_Click(object sender, EventArgs e)
        {
            CountTextBoxforForm(3);
        }

        private void materialFlatButton18_Click(object sender, EventArgs e)
        {
            CountTextBoxforForm(4);
        }

        private void monthCalendar2_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void monthCalendar2_DateSelected(object sender, DateRangeEventArgs e)
        {
            materialSingleLineTextField15.Text = monthCalendar2.SelectionStart.Date.ToShortDateString();
        }

        private void monthCalendar3_DateSelected(object sender, DateRangeEventArgs e)
        {
            materialSingleLineTextField40.Text = monthCalendar3.SelectionStart.Date.ToShortDateString();
        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }


        public void OpenPattern (int Action)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Документы (*.docx)|*.docx|Документы (*.doc)|*.doc";
            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Path = openFileDialog1.FileName;
                switch(Action)
                {
                    case 1: materialSingleLineTextField49.Text = Path; break;
                    case 2: materialSingleLineTextField50.Text = Path; break;
                    case 3: materialSingleLineTextField51.Text = Path; break;
                    case 4: materialSingleLineTextField52.Text = Path; break;
                }

            }
        }
        private void materialFlatButton20_Click(object sender, EventArgs e)
        {
            OpenPattern(1);
        }

        private void materialFlatButton21_Click(object sender, EventArgs e)
        {
            OpenPattern(2);
        }

        private void materialFlatButton22_Click(object sender, EventArgs e)
        {
            OpenPattern(3);
        }

        private void materialFlatButton23_Click(object sender, EventArgs e)
        {
            OpenPattern(4);
        }

        public int Actoin288 = 0;
        public void TextBoxesActions(int Action)
        {
            Actoin288 = 0;
            foreach (MaterialSingleLineTextField TextBox in panel1.Controls.OfType<MaterialSingleLineTextField>())
            {
                switch (Action)
                {
                    case 0: {
                                if (TextBox.Text.Length == 0)
                                {
                                    Actoin288++;
                                }
                            }; break;
                    case 1: TextBox.Text = ""; break;
                }
            }
            switch (Action)
            {
                case 0:
                    {

                                int Count = 4;
                                for (int i = 1; i <= Count; i++)
                                {
                                    switch (i)
                                    {
                                        case 1: Properties.Settings.Default.Pattern1 = materialSingleLineTextField49.Text; break;
                                        case 2: Properties.Settings.Default.Pattern2 = materialSingleLineTextField50.Text; break;
                                        case 3: Properties.Settings.Default.Pattern3 = materialSingleLineTextField51.Text; break;
                                        case 4: Properties.Settings.Default.Pattern4 = materialSingleLineTextField52.Text; break;
                                    }
                                }
                                Properties.Settings.Default.Save();
                                MessageBox.Show("Праметры успешно сохранены."); 
                        } break;
                case 1: MessageBox.Show("Праметры восстановлены по умолчанию."); break;
            }
        }

        public void RefTextBoxes(int Action)
        {

        }
        private void materialRaisedButton1_Click(object sender, EventArgs e)
        {
            TextBoxesActions(0);
        }

        private void materialFlatButton24_Click(object sender, EventArgs e)
        {
            TextBoxesActions(1);
        }

        private void materialFlatButton9_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

    }
}
