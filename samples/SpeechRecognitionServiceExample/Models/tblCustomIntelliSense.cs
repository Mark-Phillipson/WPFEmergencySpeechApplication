namespace SpeechToTextWPFSample.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("tblCustomIntelliSense")]
    public partial class CustomIntelliSense
    {
        public int ID { get; set; }

        public int? Language_ID { get; set; }

        [StringLength(255)]
        public string Display_Value { get; set; }

        public string SendKeys_Value { get; set; }

        [StringLength(255)]
        public string Command_Type { get; set; }

        public short? Category_ID { get; set; }

        [StringLength(255)]
        public string Remarks { get; set; }

        [Column(TypeName = "timestamp")]
        [MaxLength(8)]
        [Timestamp]
        public byte[] SSMA_TimeStamp { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [Required]
        public string Search { get; set; }

        public int? ComputerID { get; set; }

        [Required]
        [StringLength(30)]
        public string DeliveryType { get; set; }

        public virtual tblCategory tblCategory { get; set; }

        public virtual tblLanguage tblLanguage { get; set; }
    }
}
