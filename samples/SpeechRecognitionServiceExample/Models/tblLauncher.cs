namespace SpeechToTextWPFSample.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("tblLauncher")]
    public partial class tblLauncher
    {
        public int ID { get; set; }

        [StringLength(50)]
        public string Name { get; set; }

        [StringLength(2)]
        public string Access_Letter { get; set; }

        [StringLength(255)]
        public string CommandLine { get; set; }

        [StringLength(255)]
        public string Icon { get; set; }

        public short? Menu { get; set; }

        public int? ML_IDDepreciated { get; set; }

        public int? ProjectID { get; set; }

        public int? ComputerID { get; set; }

        public virtual tblCategory tblCategory { get; set; }
    }
}
