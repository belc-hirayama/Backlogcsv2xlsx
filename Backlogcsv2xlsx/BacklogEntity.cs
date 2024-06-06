using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Backlogcsv2xlsx
{
    public enum TableStatus
    {
        Undefined,
        Waiting,
        InDevelopment,
        Finish,
    }
    internal class BacklogEntity
    {

        [Name(Define.RequestDepartment)]
        public string Department { get; set; }  //  依頼部門

        [Name(Define.Status)]
        public string Status { get; set; }      //  状態

        [Name(Define.StartDate)]
        public string StartDate { get; set; } //  開始日

        [Name(Define.UpdateDate)]
        public string UpdateDate { get; set; } // 更新日

        [Name(Define.RakurakuExpenseNo)]
        public string RakurakuExpenseNo { get; set; }   //  楽楽精算No

        [Name(Define.CategoryName)]
        public string CategoryName { get; set; } // カテゴリー名

        private TableStatus TableHeaderStatus = TableStatus.Undefined;

        public TableStatus GetTableStatus() { return TableHeaderStatus; }
        public void SetTableStatus(TableStatus value) { TableHeaderStatus = value; }
        public void Trim()
        {
            this.RakurakuExpenseNo = this.RakurakuExpenseNo.Trim();
            this.Department = this.Department.Trim();
            this.CategoryName = this.CategoryName.Trim();
        }
    }

}
