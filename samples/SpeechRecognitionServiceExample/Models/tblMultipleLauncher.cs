namespace SpeechToTextWPFSample.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("tblMultipleLauncher")]
    public partial class tblMultipleLauncher
    {
        public int ID { get; set; }

        [StringLength(70)]
        public string Description { get; set; }
    }
}
