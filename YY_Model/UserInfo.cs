using System;
using System.ComponentModel;

namespace YY_Model
{
    /// <summary>
    /// 学生信息模型   TODO:注意：大小写和数据库保持一致
    /// </summary>
    public class UserInfo
    {
        [Description("学生编号")]
        public int? Id { get; set; }

        [Description("学生姓名")]
        public string UserName { get; set; }

        [Description("学生性别")]
        public string Sex { get; set; }

        [Description("学生联系方式")]
        public string Phone { get; set; }

        [Description("学生描述")]
        public string Description { get; set; }

        [Description("学生爱好")]
        public string Hobby { get; set; }
    }
}
