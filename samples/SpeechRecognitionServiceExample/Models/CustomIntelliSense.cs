namespace SpeechToTextWPFSample.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CustomIntelliSense")]
    public partial class CustomIntelliSense
    {
        public int ID { get; set; }

        public int? LanguageID { get; set; }

        [StringLength(255)]
        public string Display_Value { get; set; }

        public string SendKeys_Value { get; set; }

        [StringLength(255)]
        public string Command_Type { get; set; }

        public int? CategoryID { get; set; }

        [StringLength(255)]
        public string Remarks { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [Required]
        public string Search { get; set; }

        public int? ComputerID { get; set; }

        [Required]
        [StringLength(30)]
        public string DeliveryType { get; set; }

        public virtual Category Category { get; set; }

        public virtual Language Language { get; set; }
    }
}
