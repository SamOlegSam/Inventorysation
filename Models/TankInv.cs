//------------------------------------------------------------------------------
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
    using System.Collections.Generic;
    
    public partial class TankInv
    {
        public int Id { get; set; }
        public System.DateTime Data { get; set; }
        public int Filial { get; set; }
        public int Rezer { get; set; }
        public Nullable<decimal> Urov { get; set; }
        public Nullable<decimal> UrovH2O { get; set; }
        public Nullable<decimal> UrovNeft { get; set; }
        public Nullable<decimal> VNeft { get; set; }
        public Nullable<decimal> Temp { get; set; }
        public Nullable<decimal> P { get; set; }
        public Nullable<decimal> MassaBrutto { get; set; }
        public Nullable<decimal> H2O { get; set; }
        public Nullable<decimal> Salt { get; set; }
        public Nullable<decimal> Meh { get; set; }
        public Nullable<decimal> BalProc { get; set; }
        public Nullable<decimal> BalTonn { get; set; }
        public Nullable<decimal> MassaNetto { get; set; }
        public Nullable<decimal> HMim { get; set; }
        public Nullable<decimal> VMin { get; set; }
        public Nullable<decimal> MBalMin { get; set; }
        public Nullable<decimal> MNettoMin { get; set; }
        public Nullable<decimal> VTeh { get; set; }
        public Nullable<decimal> MBalTeh { get; set; }
        public Nullable<decimal> MNettoTeh { get; set; }
        public Nullable<decimal> MNettoTov { get; set; }
        public int type { get; set; }
        public decimal V { get; set; }
        public decimal VH2O { get; set; }
    }
}
