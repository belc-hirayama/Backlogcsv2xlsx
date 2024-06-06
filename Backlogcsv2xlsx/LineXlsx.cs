using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace Backlogcsv2xlsx
{
    internal class XlsxPage
    {
        public SortedDictionary<int, List<BacklogEntity>> xlsxLines;            //  キーはsortKeyに対応する数値
        Regex rakurakuRegex = new Regex(@"^[a-zA-Z]+-(?<RakurakuNo>\d+)$");     //  楽楽精算はアルファベット群-数字群
        Regex dateRegex = new Regex(@"^\d{4}/\d{2}/\d{2}\s*\d{2}:\d{2}$");      //  年月日
        DateTime beginingPeriod;
        DateTime endingPeriod;
        Dictionary<string, int> sortKey;

        public XlsxPage(DateTime begin, DateTime end, Dictionary<string, int> sortKeyDict) {
            xlsxLines = new SortedDictionary<int, List<BacklogEntity>>();
            if (begin <= end)   //  期間設定
            {
                beginingPeriod = new DateTime(begin.Year, begin.Month, begin.Day, 0,0,0); 
                endingPeriod = new DateTime(end.Year, end.Month, end.Day, 23, 59, 59);
            } else
            {
                beginingPeriod = new DateTime(end.Year, end.Month, end.Day, 0, 0, 0);
                endingPeriod = new DateTime(begin.Year, begin.Month, begin.Day, 23, 59, 59);
            }
            sortKey = sortKeyDict;
        }

        /// <summary>
        //  BacklogEntityが条件に適合した場合に辞書に格納する
        /// </summary>
        /// <param name="backlogEntity">格納するデータ</param>
        /// <param name="getDictKeyElement">backlogEntityのどの値をキーとして扱うかを決定する関数</param>
        /// <param name="multiKeyHandling">キーがカンマ区切りで複数存在した場合の扱い</param>
        public void Assignment(BacklogEntity backlogEntity, Func<BacklogEntity, string> getDictKeyElement, Func<List<string>, List<string>> multiKeyHandling)
        {
            bool flg = false;                                           //  true => backlogEntityをxlsxLinesに追加
            List<int> dictKey = new List<int>();
            if (rakurakuRegex.IsMatch(backlogEntity.RakurakuExpenseNo)) //  楽楽精算Noあり
            {
                List<string> strDictKeys = multiKeyHandling(getDictKeyElement(backlogEntity).Split(',').ToList());  //  キーを","で分割
                if (strDictKeys.Count == 0) strDictKeys.Append("");                                                 //  空白だった場合は[""]とする
                foreach (string _strDictKey in strDictKeys)
                {
                    var strDictKey = _strDictKey.Trim();
                    if (sortKey.ContainsKey(strDictKey)) dictKey.Add(sortKey[strDictKey]);                          //  キーに対応するsortedDictionary用キーを取得
                    else {dictKey.Add(sortKey.Count); sortKey[strDictKey] = sortKey.Count; }                     //  キーが辞書に登録されていなければ辞書登録
                }
                if (backlogEntity.Status != "完了")
                {
                    if (backlogEntity.StartDate.Length == 0)    // 未着手
                    {
                        backlogEntity.SetTableStatus(TableStatus.Waiting); flg = true;
                    }
                    else                                        // 開発中
                    {
                        backlogEntity.SetTableStatus(TableStatus.InDevelopment); flg = true;
                    }
                }
                else if (dateRegex.IsMatch(backlogEntity.UpdateDate))
                {
                    DateTime updateDate = DateTime.Parse(backlogEntity.UpdateDate);
                    if (updateDate >= beginingPeriod && updateDate <= endingPeriod)    //  完了
                    {
                        backlogEntity.SetTableStatus(TableStatus.Finish);
                        backlogEntity.RakurakuExpenseNo = rakurakuRegex.Match(backlogEntity.RakurakuExpenseNo).Groups["RakurakuNo"].Value;  //  楽楽精算No登録
                        flg = true;
                    }
                }
            }
            if (flg)
            {
                foreach (int intDictKey in dictKey)
                {
                    if (!xlsxLines.ContainsKey(intDictKey)) xlsxLines[intDictKey] = new List<BacklogEntity>();  //  追加する辞書キーが無ければ作成
                    xlsxLines[intDictKey].Add(backlogEntity);
                }
            }
            return;
        }

        /// <summary>
        //  格納されたデータをstring型の2次元配列として返す
        /// </summary>
        /// <param name="A1Text">辞書キーが示す項目のタイトル</param>
        /// <param name="isShowRakuraku">楽楽精算の列の有無</param>
        /// <returns>格納されたデータ</returns>
        public List<List<object>> GetxlsxLiness(string A1Text, bool isShowRakuraku)
        {
            List<List<object>> xlsxLinesList = new List<List<object>>();
            var headerText = new List<object>();
            headerText.Add(A1Text); headerText.Add("未着手"); headerText.Add("開発中"); headerText.Add("完了");
            if (isShowRakuraku) { headerText.Add("楽々精算No（完了分）"); }
            xlsxLinesList.Add(headerText);
            foreach (var line in xlsxLines)
            {
                var addLine = new List<object> {
                    sortKey.Keys.ToList()[line.Key],
                    DefaultOrNull(line.Value.Count(x => x.GetTableStatus() == TableStatus.Waiting)),
                    DefaultOrNull(line.Value.Count(x => x.GetTableStatus() == TableStatus.InDevelopment)),
                    DefaultOrNull(line.Value.Count(x => x.GetTableStatus() == TableStatus.Finish)),
                };
                var y = line.Value.Where(x => x.GetTableStatus() == TableStatus.Finish).Select(x => x.RakurakuExpenseNo);
                if (isShowRakuraku) addLine.Add("'" + GetRakurakuNo(y));
                xlsxLinesList.Add(addLine);
            }
            return xlsxLinesList;
        }

        private object DefaultOrNull(int number)
        {
            if (number == 0) return "";
            return number;
        }

        private string GetRakurakuNo(IEnumerable<string> RakurakuExpenseNo)
        {
            var joinString = new List<int>() ;
            foreach (var line in RakurakuExpenseNo)
            {
                joinString.Add(int.Parse(line));
            }
            return string.Join(",", joinString);
        }

    }
}
