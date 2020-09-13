namespace SpeechToTextWPFSample.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;
    [Table("Launcher")]
    public partial class Launcher
    {
        public int ID { get; set; }

        [StringLength(50)]
        public string Name { get; set; }

        [StringLength(2)]
        public string Access_Letter { get; set; }

        [StringLength(255)]
        public string CommandLine { get; set; }

        public int? CategoryID { get; set; }

        public int? ML_IDDepreciated { get; set; }

        public int? ProjectID { get; set; }

        public int? ComputerID { get; set; }

        public virtual Category Category { get; set; }
    }
}
