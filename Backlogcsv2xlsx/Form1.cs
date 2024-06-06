using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;
using CsvHelper;
using CsvHelper.Configuration;
using ClosedXML.Excel;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Diagnostics;
using System.Linq;

namespace Backlogcsv2xlsx
{
    public partial class Form1 : Form
    {
        private DateTime startDate;
        private DateTime endDate;
        private string sortKeyDigital;
        private string sortKeyDepartment;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "ファイル選択後に実行ボタンで実行";
            startDate = GetNearestDayOfWeek(DateTime.Now.AddDays(-7), LoadDayOfWeek2Int()); endDate = startDate.AddDays(LoadCrossMeetingSpan());    //  初期状態で選択されている日付の設定
            sortKeyDigital = ConfigurationManager.AppSettings["SortKeyDigital"];
            sortKeyDepartment = ConfigurationManager.AppSettings["SortKeyDepartment"];
            startCalendar.MaxSelectionCount = 1; endCalendar.MaxSelectionCount = 1;
            endCalendar.SetDate(endDate); _textBoxEndDate.Text = endDate.ToString("yyyy/MM/dd");
            startCalendar.SetDate(startDate); _textBoxStartDate.Text = startDate.ToString("yyyy/MM/dd");
            _textBoxDiff.KeyPress += NumberOnlyTextBox;
            _textBoxEndDate.KeyPress += DateOnlyTextBox; _textBoxStartDate.KeyPress += DateOnlyTextBox;
            if (ConfigurationManager.AppSettings["isOpenDialog"] == "1")
            {
                _Button_OpenFile.PerformClick();    //  起動時にファイル選択ダイアログを表示
                this.Activate();                    //  ファイル選択後アクティブに
            }
        }

        private DayOfWeek LoadDayOfWeek2Int()
        {
            var day = DayOfWeek.Monday;
            switch (ConfigurationManager.AppSettings["DayOfWeekEnd"])
            {
                case "日曜日":
                case "日":
                    day = DayOfWeek.Sunday;
                    break;
                case "土曜日":
                case "土":
                    day = DayOfWeek.Saturday;
                    break;
                case "金曜日":
                case "金":
                    day = DayOfWeek.Friday;
                    break;
                case "木曜日":
                case "木":
                    day = DayOfWeek.Thursday;
                    break;
                case "水曜日":
                case "水":
                    day = DayOfWeek.Wednesday;
                    break;
                case "火曜日":
                case "火":
                    day = DayOfWeek.Tuesday;
                    break;
                case "月曜日":
                case "月":
                default:
                    day =  DayOfWeek.Monday;
                    break;
            }
            return day;
        }

        private int LoadCrossMeetingSpan()
        {
           int span;
            if (int.TryParse(ConfigurationManager.AppSettings["CrossMeetingSpan"], out span) && 0 < span) return span - 1;
            return 6;
        }

        private void NumberOnlyTextBox(object sender, KeyPressEventArgs e)
        {
            if (!((e.KeyChar>= '0' && '9' >= e.KeyChar) || e.KeyChar == (char)Keys.Back || e.KeyChar == (char)Keys.Delete))
            {
                //押されたキーが 0～9でない場合は、イベントをキャンセルする
                e.Handled = true;
            }
        }

        private void DateOnlyTextBox(object sender, KeyPressEventArgs e)
        {
            if (!((e.KeyChar >= '0' && '9' >= e.KeyChar) || e.KeyChar == (char)Keys.Back || e.KeyChar == (char)Keys.Delete || (e.KeyChar == '/')))
            {
                //  押されたキーが 日付に使用する文字でない場合は、イベントをキャンセルする
                //  後で正規表現でチェックしているため不要か
                e.Handled = true;
            }
        }

        private void ReadCsv()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.Title = "開くファイルを選択してください";
            ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                _TextBox_FileName.Text = ofd.FileName;
            }
        }

        private void Csv2xlsx()
        {
            if (!File.Exists(_TextBox_FileName.Text)) { MessageBox.Show("ファイルが見つかりません。"); return; }
            this.Text = "処理中...";
            button1.Visible = false;
            try
            {
                using (var stream = new StreamReader(_TextBox_FileName.Text, Encoding.GetEncoding("Shift_JIS")))
                using (var csv = new CsvReader(stream, new CultureInfo("ja-JP", false)))
                {
                    var forCrossMeetingSortKey = File2Dict(sortKeyDepartment);
                    var forDigitalSortKey = File2Dict(sortKeyDigital);
                    var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                    {
                        HasHeaderRecord = true,
                    };
                    csv.Read();
                    csv.ReadHeader();
                    var records = csv.GetRecords<BacklogEntity>();
                    var forCrossMeeting = new XlsxPage(startCalendar.SelectionStart, endCalendar.SelectionStart, forCrossMeetingSortKey);
                    var forDigital = new XlsxPage(startCalendar.SelectionStart, endCalendar.SelectionStart, forDigitalSortKey);
                    //var forCrossMeetingGroup = records.Where(x => x.Department != "デジタル推進室").ToList();    //  ←ダメな例
                    //var forDigitalGroup = records.Where(x => x.Department == "デジタル推進室").ToList();
                    //Debug.WriteLine(forCrossMeetingGroup.Count());
                    //Debug.WriteLine(forDigitalGroup.Count());

                    foreach (var record in records)
                    {
                        record.Trim();
                        var categoryList = record.Department.Split(',').ToList();
                        for (var i = 0; i < categoryList.Count; i++) categoryList[i] = categoryList[i].Trim(); 
                        if (categoryList.Contains("デジタル推進室"))
                            forDigital.Assignment(record, r => r.CategoryName, stringList => new List<string>() { stringList[0] });
                        categoryList.RemoveAll(str => str == "デジタル推進室");
                        if (categoryList.Count > 0)
                            forCrossMeeting.Assignment(record, r => r.Department, stringList => categoryList);
                    }
                    list2Xlsx(forCrossMeeting.GetxlsxLiness("各部一覧", true), forDigital.GetxlsxLiness("デジ推一覧", false), _TextBox_FileName.Text);
                }
            } catch (Exception e)
            {
                MessageBox.Show($"Error:{e.ToString()}");
                throw (e);
            }
            button1.Visible = false;
            this.Text = "完了";
            return;
        }

        private void list2Xlsx(List<List<object>> list1, List<List<object>> list2, string fileName)
        {
            var outFileName = $"{fileName}.xlsx";
            using (XLWorkbook workbook = new XLWorkbook())
            {
                IXLWorksheet worksheet = workbook.AddWorksheet("各部一覧");
                worksheet.Cell("A1").InsertData(list1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; //  データ設定、中央寄せ
                int lastRow = worksheet.LastRowUsed().RowNumber();                                                      //  =SUM()
                worksheet.Cell(lastRow + 1, 1).Value = "合計";
                for (int i = 1; i < list1[0].Count; i++)
                {
                    if (3 < i) worksheet.Cell(lastRow + 1, i + 1).Value = " ";
                    else worksheet.Cell(lastRow + 1, i + 1).FormulaA1 = $"=SUM({(char)('A' + i)}{2}:{(char)('A' + i)}{lastRow})";
                }
                worksheet.Cells().Style.Font.FontName = "Yu Gothic";                                                    //  font : "Yu Gothic"
                worksheet.Cells().Style.Font.FontSize = 10;                                                             //  fontsize = 12
                worksheet.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;                      //  col1 : 左寄せ
                worksheet.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;                         //  row1 : 左寄せ
                worksheet.Row(1).Style.Font.FontSize = 8;                                                               //  row1 : fontsize = 8
                worksheet.RangeUsed().CreateTable().Theme = ClosedXML.Excel.XLTableTheme.TableStyleMedium2;             //  テーマ設定
                worksheet.ColumnsUsed().AdjustToContents();                                                             //  幅は自動調整
                worksheet.LastRowUsed().CellsUsed().Style.Border.TopBorder = XLBorderStyleValues.Double;                //  合計欄は二重線

                worksheet = workbook.AddWorksheet("デジ推一覧");
                worksheet.Cell("A1").InsertData(list2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                lastRow = worksheet.LastRowUsed().RowNumber();
                worksheet.Cell(lastRow + 1, 1).Value = "合計";
                for (int i = 1; i < list2[0].Count; i++)
                {
                    if (3 < i) worksheet.Cell(lastRow + 1, i + 1).Value = "";
                    else worksheet.Cell(lastRow + 1, i + 1).FormulaA1 = $"=SUM({(char)('A' + i)}{2}:{(char)('A' + i)}{lastRow})";
                }
                worksheet.Cells().Style.Font.FontName = "Yu Gothic";                                                    //  font : "Yu Gothic"
                worksheet.Cells().Style.Font.FontSize = 10;                                                             //  fontsize = 12
                worksheet.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;                      //  col1 : 左寄せ
                worksheet.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;                         //  row1 : 左寄せ
                worksheet.Row(1).Style.Font.FontSize = 8;                                                               //  row1 : fontsize = 8
                worksheet.RangeUsed().CreateTable().Theme = ClosedXML.Excel.XLTableTheme.TableStyleMedium2;             //  テーマ設定
                worksheet.ColumnsUsed().AdjustToContents();                                                             //  幅は自動調整
                worksheet.LastRowUsed().CellsUsed().Style.Border.TopBorder = XLBorderStyleValues.Double;                //  合計欄は二重線

                workbook.SaveAs(outFileName);
            }
            if (ConfigurationManager.AppSettings["isOpenXlsx"] == "1")
            {
                var startInfo = new ProcessStartInfo()
                {
                    FileName = outFileName,
                    UseShellExecute = true,
                    CreateNoWindow = true,
                };
                Process.Start(startInfo);
            } else
            {
                MessageBox.Show($"正常に出力されました。\n{outFileName}");
            }
            if (ConfigurationManager.AppSettings["isClose"] == "1")
            {
                this.Close();
            }
            return;
        }

        private void _Button_OpenFile_Click(object sender, EventArgs e)
        {
            ReadCsv();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Csv2xlsx();
        }

        //  テキストファイルからキーの並び順辞書作成
        private Dictionary<string, int> File2Dict(string fileName)
        {
            Dictionary<string, int> dict = new Dictionary<string, int>();
            if (File.Exists(fileName))
            {
                var lines = File.ReadAllLines(fileName, Encoding.GetEncoding("shift-jis"));
                int cnt = 0;
                foreach (var line in lines)
                {
                    dict[line] = cnt++;
                }
            }
            return dict;
        }

        private void endCalendar_DateChanged(object sender, DateRangeEventArgs e)
        {
            SetEndCalender(endCalendar.SelectionStart);
        }

        private void startCalendar_DateChanged(object sender, DateRangeEventArgs e)
        {
            SetStartCalender(startCalendar.SelectionStart);
        }

        /// <summary>
        //  フォーム内の開始日付を変更する
        /// </summary>
        /// <param name="dateTime">更新する日付</param>
        private void SetStartCalender(DateTime dateTime)
        {
            _textBoxDiff.TextChanged -= _textBoxDiff_TextChanged;   //  フォームを書き換えた際に他のイベントが発火しないよう一時的に削除
            startCalendar.DateChanged -= startCalendar_DateChanged;
            _textBoxStartDate.Text = dateTime.ToString("yyyy/MM/dd");
            startCalendar.SetDate(dateTime);
            CalenderChenged();
            _textBoxDiff.TextChanged += _textBoxDiff_TextChanged;
            startCalendar.DateChanged += startCalendar_DateChanged;
        }
        /// <summary>
        //  フォーム内の終了日付を変更する
        /// </summary>
        /// <param name="dateTime">更新する日付</param>

        private void SetEndCalender(DateTime dateTime)
        {
            _textBoxDiff.TextChanged -= _textBoxDiff_TextChanged;   //  フォームを書き換えた際に他のイベントが発火しないよう一時的に削除
            endCalendar.DateChanged -= endCalendar_DateChanged;
            _textBoxEndDate.Text = dateTime.ToString("yyyy/MM/dd");
            endCalendar.SetDate(dateTime);
            CalenderChenged();
            _textBoxDiff.TextChanged += _textBoxDiff_TextChanged;
            endCalendar.DateChanged += endCalendar_DateChanged;
        }

        /// <summary>
        //  day以降の最も近い(wantDayOfWeek)曜日の日を探すs
        /// </summary>
        /// <param name="day">検索の起点となる日</param>
        /// <param name="wantDayOfWeek">探す曜日</param>
        /// <remarns>(day)以降の初めての(wantDayOfWeek)曜日</returns>
        private DateTime GetNearestDayOfWeek(DateTime day, DayOfWeek wantDayOfWeek)
        {
            DayOfWeek dayOfWeek = day.DayOfWeek;
            return (dayOfWeek == wantDayOfWeek) ? day : day.AddDays((int)(DayOfWeek.Saturday - day.DayOfWeek + wantDayOfWeek) % 7 + 1);
        }

        private void _radioButtonDay_CheckedChanged(object sender, EventArgs e)
        {
            _textBoxDiff.Text = ((endCalendar.SelectionStart - startCalendar.SelectionStart).Days + 1).ToString();
        }

        private void _radioButtonWeek_CheckedChanged(object sender, EventArgs e)
        {
            _textBoxDiff.Text = (((endCalendar.SelectionStart - startCalendar.SelectionStart).Days + 1)/7).ToString();
        }

        /// <summary>
        //  カレンダー変更時、選択日数欄を更新
        /// </summary>
        private void CalenderChenged()
        {
            if (_radioButtonDay.Checked)
            {
                _textBoxDiff.Text = ((endCalendar.SelectionStart - startCalendar.SelectionStart).Days + 1).ToString();
            } else
            {
                _textBoxDiff.Text = (((endCalendar.SelectionStart - startCalendar.SelectionStart).Days + 1) / 7).ToString();
            }
        }

        private void _textBoxDiff_TextChanged(object sender, EventArgs e)
        {
            int inputNumber;
            if (int.TryParse(_textBoxDiff.Text, out inputNumber))
            {
                if (_radioButtonWeek.Checked)
                {
                    inputNumber *= 7;
                }
                SetStartCalender(endCalendar.SelectionStart.AddDays(-1 * inputNumber + 1));
            }
        }

        private void _textBoxStartDate_LostFocus(object sender, EventArgs e)
        {
            DateTime startDate;
            var text = _textBoxStartDate.Text;
            if (String2Date(text, DateTime.Now, out startDate)) SetStartCalender(startDate);
            else _textBoxStartDate.Text = startCalendar.SelectionStart.ToString("yyyy/MM/dd");
        }

        private void _textBoxEndDate_LostFocus(object sender, EventArgs e)
        {
            DateTime endDate;
            var text = _textBoxEndDate.Text;
            if (String2Date(text, DateTime.Now, out endDate)) SetEndCalender(endDate);
            else _textBoxEndDate.Text = endCalendar.SelectionStart.ToString("yyyy/MM/dd");
        }

        /// <summary>
        //  任意長の数字列もしくはyy/MM/ddから日付を返すメソッド
        //  MDなどのアレ
        /// </summary>
        //  <param name="dateString">入力文字列</param>
        //  <param name="defaultDate ">未入力だった場合に設定される年月日</param>
        //  <param name="date">変換後の日付</param>
        //  <returns>正常な日付に変換できればtrue、そうでなければfalse</returns>
        private bool String2Date(string dateString, DateTime defaultDate, out DateTime date)
        {
            if (dateString.Length == 0) //  何も入力されていなければデフォルト値を返す
            {
                date = defaultDate; return true;
            }
            if (DateTime.TryParse(dateString, out date))    //  正しくparseできた場合
            {
                return true;
            }
            else if (Regex.IsMatch(dateString, @"^\d{1,2}$") && 1 <= int.Parse(dateString) && int.Parse(dateString) < DateTime.DaysInMonth(defaultDate.Year, defaultDate.Month))
            {
                //  日付のみ入力
                date = new DateTime(defaultDate.Year, defaultDate.Month, int.Parse(dateString));
                return true;
            }
            else if (Regex.IsMatch(dateString, @"(^(?<month>\d{1,2})(?<day>\d{2})$)|(^(?<month>\d+)/(?<day>\d+)$)"))
            {
                //  月日のみ入力
                var groups = Regex.Match(dateString, @"(^(?<month>\d{1,2})(?<day>\d{2})$)|(^(?<month>\d+)/(?<day>\d+)$)").Groups;
                var month = int.Parse(groups["month"].Value);
                var day = int.Parse(groups["day"].Value);
                if (1 <= month && month <= 12 && 1 <= day && day <= DateTime.DaysInMonth(defaultDate.Year, month))
                {
                    date = new DateTime(defaultDate.Year, month, day);
                    return true;
                }
            }
            else if (Regex.IsMatch(dateString, @"(^(?<year>\d{1,4})(?<month>\d{2})(?<day>\d{2})$)|(^(?<year>\d+)/(?<month>\d+)/(?<day>\d+)$)"))
            {
                //  年月日入力
                var groups = Regex.Match(dateString, @"^((?<year>\d{1,4})(?<month>\d{2})(?<day>\d{2}))|((?<year>\d+)/(?<month>\d+)/(?<day>\d+))$").Groups;
                var year = int.Parse(groups["year"].Value);
                var month = int.Parse(groups["month"].Value);
                var day = int.Parse(groups["day"].Value);
                if (0 <= year && year <= 9) year += defaultDate.Year - defaultDate.Year % 10;
                else if (0 <= year && year <= 99) year += defaultDate.Year - defaultDate.Year % 100;
                else if (0 <= year && year <= 999) year += defaultDate.Year - defaultDate.Year % 1000;
                if (1800 <= year && year <= 2500 && 1 <= month && month <= 12 && 1 <= day && day <= DateTime.DaysInMonth(year, month))
                {
                    date = new DateTime(year, month, day);
                    return true;
                }
            }
            return false;
        }

        private void _TextBox_FileName_DragDrop(object sender, DragEventArgs e)
        {
            // ファイルが渡されていなければ、何もしない
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

            // 渡されたファイルに対して処理を行う
            _TextBox_FileName.Text = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
        }

        private void _TextBox_FileName_DragEnter(object sender, DragEventArgs e)
        {
            // ドラッグドロップ時にカーソルの形状を変更
            e.Effect = DragDropEffects.All;
        }
    }
}
