namespace SpeechToTextWPFSample.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class tblCategory
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public tblCategory()
        {
            tblCustomIntelliSenses = new HashSet<CustomIntelliSense>();
            tblLaunchers = new HashSet<tblLauncher>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public short MenuNumber { get; set; }

        [StringLength(30)]
        public string Category { get; set; }

        [StringLength(255)]
        public string Category_Type { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CustomIntelliSense> tblCustomIntelliSenses { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<tblLauncher> tblLaunchers { get; set; }
    }
}
