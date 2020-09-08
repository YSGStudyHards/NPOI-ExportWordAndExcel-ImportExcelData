using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using YY_NpoiExportAndImport.Models;

namespace YY_NpoiExportAndImport.Data
{
    /// <summary>
    /// TODO:数据库上下文类是为给定数据模型协调 EF Core 功能的主类。 上下文派生自 Microsoft.EntityFrameworkCore.DbContext。 上下文指定数据模型中包含哪些实体。
    /// 定义数据库上下文
    /// </summary>
    public class CaseInfoContext:DbContext
    {
        public CaseInfoContext(DbContextOptions<CaseInfoContext> options)
    : base(options)
        {
        }

        //创建 DbSet<TEntity> 属性。 在 EF Core 术语中：
        //实体集通常对应数据库表。
        //实体对应表中的行。

        //TODO:注意： EF 创建一系列数据表，表名默认和 DbSet 属性名相同（因为实体集合包含多个实体，因此DbSet属性名称应为复数形式）

        /// <summary>
        /// 案件信息模型
        /// </summary>
        public DbSet<CaseInfo> CaseInfos { get; set; }

        //当数据库创建完成后， EF 创建一系列数据表，表名默认和 DbSet 属性名相同。 集合属性的名称一般使用复数形式，但不同的开发人员的命名习惯可能不一样，
        //开发人员根据自己的情况确定是否使用复数形式。 在定义 DbSet 属性的代码之后，添加下面高亮代码，对 DbContext 指定单数的表名来覆盖默认的表名。

        /// <summary>
        /// 重写OnModelCreating方法，配置映射
        /// </summary>
        /// <param name="modelBuilder"></param>
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            // 配置表名映射
            modelBuilder.Entity<CaseInfo>().ToTable("CaseInfo");
            base.OnModelCreating(modelBuilder);
        }
    }
}
