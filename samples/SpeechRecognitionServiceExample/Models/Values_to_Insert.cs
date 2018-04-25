namespace SpeechToTextWPFSample.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Values to Insert")]
    public partial class Values_to_Insert
    {
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ID { get; set; }

        [Column("Value to Insert")]
        [Required]
        [StringLength(255)]
        public string Value_to_Insert { get; set; }

        [Required]
        [StringLength(255)]
        public string Lookup { get; set; }

        [StringLength(255)]
        public string Description { get; set; }

        [Column(TypeName = "timestamp")]
        [MaxLength(8)]
        [Timestamp]
        public byte[] RowVersion { get; set; }
    }
}
