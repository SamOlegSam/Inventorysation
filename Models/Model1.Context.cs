﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан по шаблону.
//
//     Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//     Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Inventariz.Models
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class asopnEntities : DbContext
    {
        public asopnEntities()
            : base("name=asopnEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<act> act { get; set; }
        public virtual DbSet<act_tank> act_tank { get; set; }
        public virtual DbSet<calibration> calibration { get; set; }
        public virtual DbSet<filials> filials { get; set; }
        public virtual DbSet<oper> oper { get; set; }
        public virtual DbSet<Podpisanty> Podpisanty { get; set; }
        public virtual DbSet<qpass> qpass { get; set; }
        public virtual DbSet<qpass_tank> qpass_tank { get; set; }
        public virtual DbSet<shift> shift { get; set; }
        public virtual DbSet<sysdiagrams> sysdiagrams { get; set; }
        public virtual DbSet<taginfo> taginfo { get; set; }
        public virtual DbSet<tags> tags { get; set; }
        public virtual DbSet<tankinfo> tankinfo { get; set; }
        public virtual DbSet<TankInv> TankInv { get; set; }
        public virtual DbSet<tanktags> tanktags { get; set; }
        public virtual DbSet<twohours> twohours { get; set; }
        public virtual DbSet<twohoursconf> twohoursconf { get; set; }
        public virtual DbSet<units> units { get; set; }
        public virtual DbSet<trl_tank> trl_tank { get; set; }
        public virtual DbSet<trl_tanktype> trl_tanktype { get; set; }
        public virtual DbSet<trl_tankview> trl_tankview { get; set; }
    }
}
