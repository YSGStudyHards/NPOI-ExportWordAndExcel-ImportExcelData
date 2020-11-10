using Microsoft.VisualBasic.CompilerServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace YY_NpoiExportAndImport.Models
{
    public class CaseInfo
    {
        /// <summary>
        /// 案件编号
        /// </summary>
        [Description("案件编号")]
        public string 案件编号 { get; set; }

        /// <summary>
        /// 案件名称
        /// </summary>
        [Description("案件名称")]
        public string 案件名称 { get; set; }

        /// <summary>
        /// 小案类别
        /// </summary>
        [Description("小案类别")]
        public string 小案类别 { get; set; }

        /// <summary>
        /// 立案单位
        /// </summary>
        [Description("立案单位")]
        public string 立案单位 { get; set; }

        /// <summary>
        /// 立案日期
        /// </summary>
        [Description("立案日期")]
        [DataType(DataType.Date)]
        public DateTime 立案日期 { get; set; }

        /// <summary>
        /// 发案时间下限
        /// </summary>
        [Description("发案时间下限")]
        [DataType(DataType.Date)]
        public DateTime 发案时间下限 { get; set; }

        /// <summary>
        /// 管辖派出所
        /// </summary>
        [Description("管辖派出所")]
        public string 管辖派出所 { get; set; }

        /// <summary>
        /// 报案人证件号码
        /// </summary>
        [Description("报案人证件号码")]
        public string 报案人证件号码 { get; set; }

        /// <summary>
        /// 报案人单位
        /// </summary>
        [Description("报案人单位")]
        public string 报案人单位 { get; set; }

        /// <summary>
        /// 发案地址
        /// </summary>
        [Description("发案地址")]
        public string 发案地址 { get; set; }
    }
}
